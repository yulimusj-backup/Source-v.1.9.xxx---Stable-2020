<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiAOP_AdjustDPP2
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
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppliersCodeNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppliersProductDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPONO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColGRNNO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStockOutNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppDONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLotNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalQtyMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitsCodeMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDesc = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColTotalQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNettPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColATCost = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColShowDetail = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRemarksItem = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryGRN = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryMasterUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReferedToPO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAtAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscFromPerct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnitsMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMasterUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyPerMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAdjustment = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAdjustmentReason = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAdjustmentUser = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAdjustmentTimestamp = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(773, 518)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(106, 32)
        Me.btnCancel.TabIndex = 7
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
        Me.btnOK.TabIndex = 6
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn4, Me.ColCodeNo, Me.ColDescription, Me.ColSuppliersCodeNumber, Me.ColSuppliersProductDesc, Me.ColPONO, Me.ColGRNNO, Me.ColStockOutNo, Me.ColSuppDONo, Me.ColLotNumber, Me.ColExpiry, Me.ColTotalQtyMulti, Me.ColUnitsCodeMulti, Me.ColRatioDesc, Me.ColTotalQty, Me.ColUnit, Me.ColLocation, Me.ColCurrency, Me.ColUnitPrice, Me.ColDiscount, Me.ColNettPrice, Me.ColAmount, Me.ColATCost, Me.ColShowDetail, Me.ColRemarksItem, Me.ColPrimaryCodeNo, Me.ColPrimaryPO, Me.ColPrimaryGRN, Me.ColPrimaryMasterUnits, Me.ColPrimaryUnit, Me.ColPrimarySecUnit, Me.ColPrimaryCurrency, Me.ColPrimaryLocation, Me.ColReferedToPO, Me.ColAtAmount, Me.ColDiscAmount, Me.ColDiscFromPerct, Me.ColRatioReference, Me.ColRatioDefault, Me.ColPrimaryUnitsMulti, Me.ColMasterUnits, Me.ColQtyMaster, Me.ColQtyPerMaster, Me.ColQty, Me.ColSecQty, Me.ColSecUnit, Me.ColItemAdjustment, Me.ColItemAdjustmentReason, Me.ColItemAdjustmentUser, Me.ColItemAdjustmentTimestamp})
        Me.dgvItem.Location = New System.Drawing.Point(13, 13)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(866, 497)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 8
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.Frozen = True
        Me.DataGridViewTextBoxColumn4.HeaderText = "No"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.ReadOnly = True
        Me.DataGridViewTextBoxColumn4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.DataGridViewTextBoxColumn4.Width = 32
        '
        'ColCodeNo
        '
        Me.ColCodeNo.Frozen = True
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.ReadOnly = True
        Me.ColCodeNo.Width = 88
        '
        'ColDescription
        '
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.ReadOnly = True
        Me.ColDescription.Width = 104
        '
        'ColSuppliersCodeNumber
        '
        Me.ColSuppliersCodeNumber.HeaderText = "Supplier's Code Number"
        Me.ColSuppliersCodeNumber.Name = "ColSuppliersCodeNumber"
        Me.ColSuppliersCodeNumber.ReadOnly = True
        Me.ColSuppliersCodeNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSuppliersCodeNumber.Visible = False
        Me.ColSuppliersCodeNumber.Width = 167
        '
        'ColSuppliersProductDesc
        '
        Me.ColSuppliersProductDesc.HeaderText = "Supplier's Product Desc"
        Me.ColSuppliersProductDesc.Name = "ColSuppliersProductDesc"
        Me.ColSuppliersProductDesc.ReadOnly = True
        Me.ColSuppliersProductDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSuppliersProductDesc.Visible = False
        Me.ColSuppliersProductDesc.Width = 165
        '
        'ColPONO
        '
        Me.ColPONO.HeaderText = "PO. No"
        Me.ColPONO.Name = "ColPONO"
        Me.ColPONO.ReadOnly = True
        Me.ColPONO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPONO.Visible = False
        Me.ColPONO.Width = 60
        '
        'ColGRNNO
        '
        Me.ColGRNNO.HeaderText = "GRN No"
        Me.ColGRNNO.Name = "ColGRNNO"
        Me.ColGRNNO.ReadOnly = True
        Me.ColGRNNO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColGRNNO.Visible = False
        Me.ColGRNNO.Width = 67
        '
        'ColStockOutNo
        '
        Me.ColStockOutNo.HeaderText = "Stock In No"
        Me.ColStockOutNo.Name = "ColStockOutNo"
        Me.ColStockOutNo.ReadOnly = True
        Me.ColStockOutNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColStockOutNo.Visible = False
        Me.ColStockOutNo.Width = 86
        '
        'ColSuppDONo
        '
        Me.ColSuppDONo.HeaderText = "Supplier DO No"
        Me.ColSuppDONo.Name = "ColSuppDONo"
        Me.ColSuppDONo.ReadOnly = True
        Me.ColSuppDONo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSuppDONo.Visible = False
        Me.ColSuppDONo.Width = 113
        '
        'ColLotNumber
        '
        Me.ColLotNumber.HeaderText = "Batch Number"
        Me.ColLotNumber.Name = "ColLotNumber"
        Me.ColLotNumber.ReadOnly = True
        Me.ColLotNumber.Width = 123
        '
        'ColExpiry
        '
        Me.ColExpiry.HeaderText = "Expiry Date"
        Me.ColExpiry.Name = "ColExpiry"
        Me.ColExpiry.ReadOnly = True
        Me.ColExpiry.Width = 105
        '
        'ColTotalQtyMulti
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotalQtyMulti.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColTotalQtyMulti.HeaderText = "Total Qty"
        Me.ColTotalQtyMulti.Name = "ColTotalQtyMulti"
        Me.ColTotalQtyMulti.ReadOnly = True
        Me.ColTotalQtyMulti.Width = 91
        '
        'ColUnitsCodeMulti
        '
        Me.ColUnitsCodeMulti.HeaderText = "Units"
        Me.ColUnitsCodeMulti.Name = "ColUnitsCodeMulti"
        Me.ColUnitsCodeMulti.ReadOnly = True
        Me.ColUnitsCodeMulti.Width = 65
        '
        'ColRatioDesc
        '
        Me.ColRatioDesc.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColRatioDesc.HeaderText = "Ratio"
        Me.ColRatioDesc.Name = "ColRatioDesc"
        Me.ColRatioDesc.ReadOnly = True
        Me.ColRatioDesc.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRatioDesc.Width = 47
        '
        'ColTotalQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotalQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColTotalQty.HeaderText = "Qty Default"
        Me.ColTotalQty.Name = "ColTotalQty"
        Me.ColTotalQty.ReadOnly = True
        Me.ColTotalQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTotalQty.Width = 85
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Units Default"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.ReadOnly = True
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnit.Width = 95
        '
        'ColLocation
        '
        Me.ColLocation.HeaderText = "Location"
        Me.ColLocation.Name = "ColLocation"
        Me.ColLocation.ReadOnly = True
        Me.ColLocation.Width = 87
        '
        'ColCurrency
        '
        Me.ColCurrency.HeaderText = "Currency"
        Me.ColCurrency.Name = "ColCurrency"
        Me.ColCurrency.ReadOnly = True
        Me.ColCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCurrency.Width = 71
        '
        'ColUnitPrice
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.ReadOnly = True
        Me.ColUnitPrice.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitPrice.Width = 75
        '
        'ColDiscount
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscount.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColDiscount.HeaderText = "Discount (%)"
        Me.ColDiscount.Name = "ColDiscount"
        Me.ColDiscount.ReadOnly = True
        Me.ColDiscount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDiscount.Width = 95
        '
        'ColNettPrice
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColNettPrice.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColNettPrice.HeaderText = "Nett Price"
        Me.ColNettPrice.Name = "ColNettPrice"
        Me.ColNettPrice.ReadOnly = True
        Me.ColNettPrice.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNettPrice.Width = 76
        '
        'ColAmount
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.ReadOnly = True
        Me.ColAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAmount.Width = 62
        '
        'ColATCost
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColATCost.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColATCost.HeaderText = "At Cost (IDR)"
        Me.ColATCost.Name = "ColATCost"
        Me.ColATCost.ReadOnly = True
        Me.ColATCost.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColATCost.Width = 96
        '
        'ColShowDetail
        '
        Me.ColShowDetail.HeaderText = "Detail Calculation"
        Me.ColShowDetail.Name = "ColShowDetail"
        Me.ColShowDetail.ReadOnly = True
        Me.ColShowDetail.Width = 123
        '
        'ColRemarksItem
        '
        Me.ColRemarksItem.HeaderText = "Remarks"
        Me.ColRemarksItem.Name = "ColRemarksItem"
        Me.ColRemarksItem.ReadOnly = True
        Me.ColRemarksItem.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarksItem.Width = 70
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.ReadOnly = True
        Me.ColPrimaryCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 113
        '
        'ColPrimaryPO
        '
        Me.ColPrimaryPO.HeaderText = "Primary PO"
        Me.ColPrimaryPO.Name = "ColPrimaryPO"
        Me.ColPrimaryPO.ReadOnly = True
        Me.ColPrimaryPO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryPO.Visible = False
        Me.ColPrimaryPO.Width = 86
        '
        'ColPrimaryGRN
        '
        Me.ColPrimaryGRN.HeaderText = "PrimaryGRN"
        Me.ColPrimaryGRN.Name = "ColPrimaryGRN"
        Me.ColPrimaryGRN.ReadOnly = True
        Me.ColPrimaryGRN.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryGRN.Visible = False
        Me.ColPrimaryGRN.Width = 93
        '
        'ColPrimaryMasterUnits
        '
        Me.ColPrimaryMasterUnits.HeaderText = "Primary Master Unit"
        Me.ColPrimaryMasterUnits.Name = "ColPrimaryMasterUnits"
        Me.ColPrimaryMasterUnits.ReadOnly = True
        Me.ColPrimaryMasterUnits.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryMasterUnits.Visible = False
        Me.ColPrimaryMasterUnits.Width = 138
        '
        'ColPrimaryUnit
        '
        Me.ColPrimaryUnit.HeaderText = "Primary Units"
        Me.ColPrimaryUnit.Name = "ColPrimaryUnit"
        Me.ColPrimaryUnit.ReadOnly = True
        Me.ColPrimaryUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryUnit.Visible = False
        Me.ColPrimaryUnit.Width = 98
        '
        'ColPrimarySecUnit
        '
        Me.ColPrimarySecUnit.HeaderText = "Primary Sec Unit"
        Me.ColPrimarySecUnit.Name = "ColPrimarySecUnit"
        Me.ColPrimarySecUnit.ReadOnly = True
        Me.ColPrimarySecUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimarySecUnit.Visible = False
        Me.ColPrimarySecUnit.Width = 119
        '
        'ColPrimaryCurrency
        '
        Me.ColPrimaryCurrency.HeaderText = "PrimaryCurrency"
        Me.ColPrimaryCurrency.Name = "ColPrimaryCurrency"
        Me.ColPrimaryCurrency.ReadOnly = True
        Me.ColPrimaryCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryCurrency.Visible = False
        Me.ColPrimaryCurrency.Width = 119
        '
        'ColPrimaryLocation
        '
        Me.ColPrimaryLocation.HeaderText = "Primary Location"
        Me.ColPrimaryLocation.Name = "ColPrimaryLocation"
        Me.ColPrimaryLocation.ReadOnly = True
        Me.ColPrimaryLocation.Visible = False
        Me.ColPrimaryLocation.Width = 139
        '
        'ColReferedToPO
        '
        Me.ColReferedToPO.HeaderText = "Column1"
        Me.ColReferedToPO.Name = "ColReferedToPO"
        Me.ColReferedToPO.ReadOnly = True
        Me.ColReferedToPO.Visible = False
        Me.ColReferedToPO.Width = 88
        '
        'ColAtAmount
        '
        Me.ColAtAmount.HeaderText = "At Amount"
        Me.ColAtAmount.Name = "ColAtAmount"
        Me.ColAtAmount.ReadOnly = True
        Me.ColAtAmount.Visible = False
        Me.ColAtAmount.Width = 98
        '
        'ColDiscAmount
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscAmount.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColDiscAmount.HeaderText = "Disc. Amount"
        Me.ColDiscAmount.Name = "ColDiscAmount"
        Me.ColDiscAmount.ReadOnly = True
        Me.ColDiscAmount.Visible = False
        Me.ColDiscAmount.Width = 116
        '
        'ColDiscFromPerct
        '
        Me.ColDiscFromPerct.HeaderText = "Disc From Perct"
        Me.ColDiscFromPerct.Name = "ColDiscFromPerct"
        Me.ColDiscFromPerct.ReadOnly = True
        Me.ColDiscFromPerct.Visible = False
        Me.ColDiscFromPerct.Width = 133
        '
        'ColRatioReference
        '
        Me.ColRatioReference.HeaderText = "Ratio Reference"
        Me.ColRatioReference.Name = "ColRatioReference"
        Me.ColRatioReference.ReadOnly = True
        Me.ColRatioReference.Visible = False
        Me.ColRatioReference.Width = 136
        '
        'ColRatioDefault
        '
        Me.ColRatioDefault.HeaderText = "Ratio Default"
        Me.ColRatioDefault.Name = "ColRatioDefault"
        Me.ColRatioDefault.ReadOnly = True
        Me.ColRatioDefault.Visible = False
        Me.ColRatioDefault.Width = 115
        '
        'ColPrimaryUnitsMulti
        '
        Me.ColPrimaryUnitsMulti.HeaderText = "Primary Units Multi"
        Me.ColPrimaryUnitsMulti.Name = "ColPrimaryUnitsMulti"
        Me.ColPrimaryUnitsMulti.ReadOnly = True
        Me.ColPrimaryUnitsMulti.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryUnitsMulti.Visible = False
        Me.ColPrimaryUnitsMulti.Width = 131
        '
        'ColMasterUnits
        '
        Me.ColMasterUnits.HeaderText = "Master Unit"
        Me.ColMasterUnits.Name = "ColMasterUnits"
        Me.ColMasterUnits.ReadOnly = True
        Me.ColMasterUnits.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColMasterUnits.Visible = False
        Me.ColMasterUnits.Width = 86
        '
        'ColQtyMaster
        '
        Me.ColQtyMaster.HeaderText = "Qty Master"
        Me.ColQtyMaster.Name = "ColQtyMaster"
        Me.ColQtyMaster.ReadOnly = True
        Me.ColQtyMaster.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQtyMaster.Visible = False
        Me.ColQtyMaster.Width = 83
        '
        'ColQtyPerMaster
        '
        Me.ColQtyPerMaster.HeaderText = "Qty Per Master"
        Me.ColQtyPerMaster.Name = "ColQtyPerMaster"
        Me.ColQtyPerMaster.ReadOnly = True
        Me.ColQtyPerMaster.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQtyPerMaster.Visible = False
        Me.ColQtyPerMaster.Width = 109
        '
        'ColQty
        '
        Me.ColQty.HeaderText = "Qty "
        Me.ColQty.Name = "ColQty"
        Me.ColQty.ReadOnly = True
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Visible = False
        Me.ColQty.Width = 36
        '
        'ColSecQty
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSecQty.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColSecQty.HeaderText = "@"
        Me.ColSecQty.Name = "ColSecQty"
        Me.ColSecQty.ReadOnly = True
        Me.ColSecQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSecQty.Visible = False
        Me.ColSecQty.Width = 28
        '
        'ColSecUnit
        '
        Me.ColSecUnit.HeaderText = "@ Unit"
        Me.ColSecUnit.Name = "ColSecUnit"
        Me.ColSecUnit.ReadOnly = True
        Me.ColSecUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSecUnit.Visible = False
        Me.ColSecUnit.Width = 57
        '
        'ColItemAdjustment
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColItemAdjustment.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColItemAdjustment.HeaderText = "Adjustment"
        Me.ColItemAdjustment.Name = "ColItemAdjustment"
        Me.ColItemAdjustment.ReadOnly = True
        Me.ColItemAdjustment.Visible = False
        Me.ColItemAdjustment.Width = 103
        '
        'ColItemAdjustmentReason
        '
        Me.ColItemAdjustmentReason.HeaderText = "Adjust Reason"
        Me.ColItemAdjustmentReason.Name = "ColItemAdjustmentReason"
        Me.ColItemAdjustmentReason.ReadOnly = True
        Me.ColItemAdjustmentReason.Visible = False
        Me.ColItemAdjustmentReason.Width = 125
        '
        'ColItemAdjustmentUser
        '
        Me.ColItemAdjustmentUser.HeaderText = "Adjust By"
        Me.ColItemAdjustmentUser.Name = "ColItemAdjustmentUser"
        Me.ColItemAdjustmentUser.ReadOnly = True
        Me.ColItemAdjustmentUser.Visible = False
        Me.ColItemAdjustmentUser.Width = 92
        '
        'ColItemAdjustmentTimestamp
        '
        Me.ColItemAdjustmentTimestamp.HeaderText = "Adjust Timestamp"
        Me.ColItemAdjustmentTimestamp.Name = "ColItemAdjustmentTimestamp"
        Me.ColItemAdjustmentTimestamp.ReadOnly = True
        Me.ColItemAdjustmentTimestamp.Visible = False
        Me.ColItemAdjustmentTimestamp.Width = 145
        '
        'FrmTransaksiAOP_AdjustDPP2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(892, 563)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(900, 600)
        Me.Name = "FrmTransaksiAOP_AdjustDPP2"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Price Adjustment"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Public WithEvents btnOK As System.Windows.Forms.Button
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppliersCodeNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppliersProductDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPONO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColGRNNO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStockOutNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppDONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLotNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalQtyMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitsCodeMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDesc As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColTotalQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNettPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColATCost As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColShowDetail As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRemarksItem As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryGRN As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryMasterUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReferedToPO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAtAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscFromPerct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnitsMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMasterUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyPerMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAdjustment As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAdjustmentReason As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAdjustmentUser As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAdjustmentTimestamp As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
