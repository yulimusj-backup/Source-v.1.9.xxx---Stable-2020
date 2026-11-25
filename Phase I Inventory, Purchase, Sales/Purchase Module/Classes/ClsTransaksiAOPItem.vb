Option Explicit On
Option Strict Off

Public Class ClsTransaksiAOPItem
    Public Sub IsiTBLProductByPO(ByVal vPrimaryAOP As Double, ByVal vTBL As _
   DseTransaksiAOP.tbltrans_aop_item_view1DataTable, ByVal vPrimaryPO As Double)
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_item_view1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryAOP_PrimaryPO(vTBL, vPrimaryAOP, vPrimaryPO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLPO2(ByVal vPrimaryAOP As Double, ByVal vTBL As _
      DseTransaksiAOP.tbltrans_aop_item_view1DataTable)
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_item_view1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryTransaksiAOP(vTBL, vPrimaryAOP)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function IsiTBLAOP(ByVal TBL As DseTransaksiAOP.tbltrans_aop_item_view0DataTable, _
        ByVal vTanggalSampai As Date, ByVal vTanggalDari As Date) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_item_view0TableAdapter
        DTA.ClearBeforeFill = True

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vTanggalSampai, vTanggalDari)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vPrimaryAOP As Double, ByVal vPrimaryCodeNo As Double, ByVal vPrimaryPO As Double) As Decimal
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml2(vPrimaryAOP, vPrimaryCodeNo, vPrimaryPO).HasValue = True Then
                Return DTA.ScalarQuery_GetJml2(vPrimaryAOP, vPrimaryCodeNo, vPrimaryPO).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml2(ByVal vPrimaryAOP As Double, ByVal vPrimaryCodeNo As Double, ByVal vPrimaryLocation As Double, ByVal vBatchNumber As String) As Decimal
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml3(vPrimaryAOP, vPrimaryCodeNo, vPrimaryLocation, vBatchNumber).HasValue = True Then
                Return DTA.ScalarQuery_GetJml3(vPrimaryAOP, vPrimaryCodeNo, vPrimaryLocation, vBatchNumber).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vPrimaryAOP As Double, ByVal vPrimaryLocation As Double, ByVal vPrimaryProduct As Double, _
    ByVal vBatchNumber As String) As Decimal
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vPrimaryLocation, vBatchNumber, vPrimaryProduct, vPrimaryAOP).HasValue = True Then
                Return DTA.ScalarQuery_GetJml(vPrimaryLocation, vBatchNumber, vPrimaryProduct, vPrimaryAOP).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''dipakai untuk Class GRN
    Public Function GetJmlByGRN(ByVal PrimaryGRN As Double) As Integer
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJmlByPrimaryGRN(PrimaryGRN).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiTBLGRN(ByVal vPrimaryAOP As Double, ByVal TBL As DseTransaksiAOP.DataTable9DataTable) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.DataTable9TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryAOP)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Public Function Simpan_Default(ByVal vdgvItem As DataGridView, ByVal vDgvExchangeRate As Windows.Forms.DataGridView, _
    '    ByVal vDGVDP As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vDGVPR As DataGridView, _
    '    ByVal vDGVCostItem As DataGridView, ByVal vFrm As FrmTransaksiAOP2) As Boolean
    '    Dim row As DataGridViewRow
    '    Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemTableAdapter
    '    Dim DTA1 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item_exchangerateTableAdapter
    '    Dim DTA2 As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemdp1TableAdapter
    '    Dim DTA3 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item_purchasereturnTableAdapter
    '    Dim DTA4 As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemcostTableAdapter
    '    Dim sec_qty As Double = -1
    '    Dim Objpo As New ClsTransaksiPO

    '    Dim DTA5 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item1TableAdapter
    '    Dim TBL5 As New DseTransaksiAOP.tbltrans_aop_item1DataTable
    '    Dim row5 As DseTransaksiAOP.tbltrans_aop_item1Row

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
    '        DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
    '        DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString
    '        DTA4.Connection.ConnectionString = DTA.Connection.ConnectionString
    '        DTA5.Connection.ConnectionString = DTA.Connection.ConnectionString

    '        'simpan data item
    '        For Each row In vdgvItem.Rows
    '            If IsNumeric(row.Cells(vFrm.ColSecQty.Index).Value) = False Then
    '                sec_qty = -1
    '            Else
    '                sec_qty = row.Cells(vFrm.ColSecQty.Index).Value
    '            End If

    '            'simpan ke basis data item
    '            DTA.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryCodeNo.Index).Value, row.Cells(vFrm.ColPrimaryGRN.Index).Value, row.Cells(vFrm.ColPrimaryPO.Index).Value, _
    '            row.Cells(vFrm.ColSuppDONo.Index).Value, 0, -1, row.Cells(vFrm.ColQty.Index).Value, row.Cells(vFrm.ColTotalQty.Index).Value, _
    '            row.Cells(vFrm.ColPrimaryUnit.Index).Value, sec_qty, row.Cells(vFrm.ColPrimarySecUnit.Index).Value, row.Cells(vFrm.ColUnitPrice.Index).Value, _
    '            row.Cells(vFrm.ColDiscount.Index).Value, row.Cells(vFrm.ColNettPrice.Index).Value, row.Cells(vFrm.ColAmount.Index).Value, _
    '            row.Cells(vFrm.ColATCost.Index).Value, row.Cells(vFrm.ColRemarks.Index).Value, -1, row.Cells(vFrm.ColBatchNo.Index).Value, #1/1/1950#, 0, _
    '            CDec(row.Cells(vFrm.Col_AtAmount.Index).Value), row.Cells(vFrm.ColDiscAmount.Index).Value, row.Cells(vFrm.ColDiscFromPerct.Index).Value)
    '        Next

    '        'simpan data item exchange rate
    '        For Each row In vDgvExchangeRate.Rows
    '            DTA1.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryDokumen.Index).Value, row.Cells(vFrm.ColPrimaryCurrency1.Index).Value, _
    '               row.Cells(vFrm.ColExchangeRate.Index).Value, row.Cells(vFrm.ColExcR_Remarks.Index).Value)
    '        Next

    '        'simpan item DP
    '        For Each row In vDGVDP.Rows
    '            DTA2.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColDP_PrimaryAOPDP.Index).Value, row.Cells(vFrm.ColDP_Remarks.Index).Value)
    '        Next

    '        'simpan item sales return
    '        For Each row In vDGVPR.Rows
    '            DTA3.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryPR.Index).Value, row.Cells(vFrm.ColRemarks_PurchaseReturn.Index).Value)
    '        Next

    '        'simpan item cost item
    '        For Each row In vDGVCostItem.Rows
    '            If row.Index <> vDGVCostItem.Rows.Count - 1 Then
    '                DTA4.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColDescCost.Index).Value, row.Cells(vFrm.ColPrimaryCurrencyCost.Index).Value, row.Cells(vFrm.ColExcRateCost.Index).Value, _
    '                row.Cells(vFrm.ColCostNominal.Index).Value, row.Cells(vFrm.ColRemarksCost.Index).Value, row.Cells(vFrm.ColPrimaryAccount.Index).Value)
    '            End If
    '        Next

    '        'simpan primary aop pada tbl grn
    '        DTA5.Fill(TBL5, vPrimaryTransaksi)
    '        For Each row5 In TBL5.Rows
    '            DTA5.UpdateQuery_PrimaryAOP(row5.Primary_AOP, row5.Primary_GRN)
    '        Next

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '        DTA1.Dispose()
    '        DTA2.Dispose()
    '        DTA4.Dispose()
    '        DTA3.Dispose()
    '        DTA5.Dispose()
    '        TBL5.Dispose()
    '    End Try
    'End Function

    ''@rosadinnn
    ''@MUaT
    Public Function Simpan_Multi(ByVal vdgvItem As DataGridView, ByVal vDgvExchangeRate As Windows.Forms.DataGridView, _
            ByVal vDGVDP As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vDGVPR As DataGridView, _
            ByVal vDGVCostItem As DataGridView, ByVal vFrm As FrmTransaksiAOP2) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemTableAdapter
        Dim DTA1 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item_exchangerateTableAdapter
        Dim DTA2 As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemdp1TableAdapter
        Dim DTA3 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item_purchasereturnTableAdapter
        Dim DTA4 As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemcostTableAdapter
        Dim sec_qty As Double = -1
        Dim Objpo As New ClsTransaksiPO

        Dim DTA5 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item1TableAdapter
        Dim TBL5 As New DseTransaksiAOP.tbltrans_aop_item1DataTable
        Dim row5 As DseTransaksiAOP.tbltrans_aop_item1Row
        Dim AdjTimestamp As Date = #1/1/1950#

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA4.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA5.Connection.ConnectionString = DTA.Connection.ConnectionString

            Dim color1 As Color = Core_Module.GetSMESSBackColorAll

            'simpan data item
            For Each row In vdgvItem.Rows
                If IsNumeric(row.Cells(vFrm.ColSecQty.Index).Value) = False Then
                    sec_qty = -1
                Else
                    sec_qty = row.Cells(vFrm.ColSecQty.Index).Value
                End If

                'simpan ke basis data item
                If IsDate(row.Cells(vFrm.ColItemAdjustmentTimestamp.Index).Value) = False Then
                    AdjTimestamp = #1/1/1950#
                Else
                    AdjTimestamp = CDate(row.Cells(vFrm.ColItemAdjustmentTimestamp.Index).Value)
                End If

                DTA.Insert_Multi(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryCodeNo.Index).Value, row.Cells(vFrm.ColPrimaryGRN.Index).Value, row.Cells(vFrm.ColPrimaryPO.Index).Value, _
                row.Cells(vFrm.ColSuppDONo.Index).Value, 0, -1, row.Cells(vFrm.ColQty.Index).Value, row.Cells(vFrm.ColTotalQty.Index).Value, _
                row.Cells(vFrm.ColPrimaryUnit.Index).Value, sec_qty, row.Cells(vFrm.ColPrimarySecUnit.Index).Value, row.Cells(vFrm.ColUnitPrice.Index).Value, _
                row.Cells(vFrm.ColDiscount.Index).Value, row.Cells(vFrm.ColNettPrice.Index).Value, row.Cells(vFrm.ColAmount.Index).Value, _
                row.Cells(vFrm.ColATCost.Index).Value, row.Cells(vFrm.ColRemarks.Index).Value, -1, row.Cells(vFrm.ColBatchNo.Index).Value, #1/1/1950#, 0, _
                CDec(row.Cells(vFrm.Col_AtAmount.Index).Value), row.Cells(vFrm.ColDiscAmount.Index).Value, row.Cells(vFrm.ColDiscFromPerct.Index).Value, _
                CDec(row.Cells(vFrm.ColQtyMulti_Item.Index).Value), CDbl(row.Cells(vFrm.ColPrimaryUnitsMulti_Item.Index).Value), _
                CDec(row.Cells(vFrm.ColRatioReference_Item.Index).Value), CDec(row.Cells(vFrm.ColRatioDefault_Item.Index).Value), _
                CDec(row.Cells(vFrm.ColItem_Adjustment.Index).Value), row.Cells(vFrm.ColItem_AdjustUser.Index).Value, _
                row.Cells(vFrm.ColItem_AdjustmentReason.Index).Value, AdjTimestamp, row.Cells(vFrm.ColPriceByMultiUnit.Index).Value)
            Next

            'simpan data item exchange rate
            For Each row In vDgvExchangeRate.Rows
                DTA1.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryDokumen.Index).Value, row.Cells(vFrm.ColPrimaryCurrency1.Index).Value, _
                   row.Cells(vFrm.ColExchangeRate.Index).Value, row.Cells(vFrm.ColExcR_Remarks.Index).Value)
            Next

            'simpan item DP
            'aul add
            For Each row In vDGVDP.Rows
                DTA2.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColDP_PrimaryAOPDP.Index).Value, row.Cells(vFrm.ColDP_Remarks.Index).Value, row.Cells(vFrm.ColDPTipeDP.Index).Value)
            Next

            'simpan item sales return
            For Each row In vDGVPR.Rows
                DTA3.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryPR.Index).Value, row.Cells(vFrm.ColRemarks_PurchaseReturn.Index).Value)
            Next

            'simpan item cost item
            For Each row In vDGVCostItem.Rows
                If row.Index <> vDGVCostItem.Rows.Count - 1 Then
                    DTA4.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColDescCost.Index).Value, row.Cells(vFrm.ColPrimaryCurrencyCost.Index).Value, row.Cells(vFrm.ColExcRateCost.Index).Value, _
                    row.Cells(vFrm.ColCostNominal.Index).Value, row.Cells(vFrm.ColRemarksCost.Index).Value, row.Cells(vFrm.ColPrimaryAccount.Index).Value)
                End If
            Next

            'simpan primary aop pada tbl grn
            DTA5.Fill(TBL5, vPrimaryTransaksi)
            For Each row5 In TBL5.Rows
                DTA5.UpdateQuery_PrimaryAOP(row5.Primary_AOP, row5.Primary_GRN)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            DTA2.Dispose()
            DTA4.Dispose()
            DTA3.Dispose()
            DTA5.Dispose()
            TBL5.Dispose()
        End Try
    End Function

    'Public Function Simpan_Default(ByVal vdgvItem As DataGridView, ByVal vDgvExchangeRate As Windows.Forms.DataGridView, _
    '           ByVal vDGVDP As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vDGVPR As DataGridView, _
    '           ByVal vDGVCostItem As DataGridView, ByVal vFrm As FrmTransaksiAOP_Quick2) As Boolean
    '    Dim row As DataGridViewRow
    '    Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemTableAdapter
    '    Dim DTA1 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item_exchangerateTableAdapter
    '    Dim DTA2 As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemdp1TableAdapter
    '    Dim DTA3 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item_purchasereturnTableAdapter
    '    Dim DTA4 As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemcostTableAdapter
    '    Dim Expiry As Date

    '    Dim DTA5 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item1TableAdapter
    '    Dim TBL5 As New DseTransaksiAOP.tbltrans_aop_item1DataTable
    '    Dim row5 As DseTransaksiAOP.tbltrans_aop_item1Row

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
    '        DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
    '        DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString
    '        DTA4.Connection.ConnectionString = DTA.Connection.ConnectionString
    '        DTA5.Connection.ConnectionString = DTA.Connection.ConnectionString

    '        'simpan data item
    '        For Each row In vdgvItem.Rows
    '            If IsDate(row.Cells(vFrm.ColExpiry.Index).Value) = True Then
    '                Expiry = row.Cells(vFrm.ColExpiry.Index).Value
    '            Else
    '                Expiry = #1/1/1950#
    '            End If
    '            'simpan ke basis data item
    '            DTA.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryCodeNo.Index).Value, row.Cells(vFrm.ColPrimaryGRN.Index).Value, _
    '            row.Cells(vFrm.ColPrimaryPO.Index).Value, row.Cells(vFrm.ColSuppDONo.Index).Value, 0, -1, 0, row.Cells(vFrm.ColTotalQty.Index).Value, _
    '            row.Cells(vFrm.ColPrimaryUnit.Index).Value, -1, -1, row.Cells(vFrm.ColUnitPrice.Index).Value, row.Cells(vFrm.ColDiscount.Index).Value, _
    '            row.Cells(vFrm.ColNettPrice.Index).Value, row.Cells(vFrm.ColAmount.Index).Value, row.Cells(vFrm.ColATCost.Index).Value, _
    '            row.Cells(vFrm.ColRemarksItem.Index).Value, row.Cells(vFrm.ColPrimaryLocation.Index).Value, row.Cells(vFrm.ColLotNumber.Index).Value, _
    '            Expiry, row.Cells(vFrm.ColReferedToPO.Index).Value, row.Cells(vFrm.ColAtAmount.Index).Value, row.Cells(vFrm.ColDiscAmount.Index).Value, _
    '            row.Cells(vFrm.ColDiscFromPerct.Index).Value)
    '        Next

    '        'simpan data item exchange rate
    '        For Each row In vDgvExchangeRate.Rows
    '            DTA1.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryDokumen.Index).Value, row.Cells(vFrm.ColPrimaryCurrency1.Index).Value, _
    '               row.Cells(vFrm.ColExchangeRate.Index).Value, row.Cells(vFrm.ColExcR_Remarks.Index).Value)
    '        Next

    '        'simpan item DP
    '        For Each row In vDGVDP.Rows
    '            DTA2.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryAOPDP.Index).Value, row.Cells(vFrm.ColDP_Remarks.Index).Value)
    '        Next

    '        'simpan item sales return
    '        For Each row In vDGVPR.Rows
    '            DTA3.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryPR.Index).Value, row.Cells(vFrm.ColRemarks_PurchaseReturn.Index).Value)
    '        Next

    '        'simpan item cost item
    '        For Each row In vDGVCostItem.Rows
    '            If row.Index <> vDGVCostItem.Rows.Count - 1 Then
    '                DTA4.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColDescCost.Index).Value, row.Cells(vFrm.ColPrimaryCurrencyCost.Index).Value, row.Cells(vFrm.ColExcRateCost.Index).Value, _
    '                row.Cells(vFrm.ColCostNominal.Index).Value, row.Cells(vFrm.ColRemarksCost.Index).Value, row.Cells(vFrm.ColPrimaryAccount.Index).Value)
    '            End If
    '        Next

    '        'simpan primary aop pada tbl grn
    '        DTA5.Fill(TBL5, vPrimaryTransaksi)
    '        For Each row5 In TBL5.Rows
    '            DTA5.UpdateQuery_PrimaryAOP(row5.Primary_AOP, row5.Primary_GRN)
    '        Next

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '        DTA1.Dispose()
    '        DTA2.Dispose()
    '        DTA4.Dispose()
    '        DTA3.Dispose()
    '        DTA5.Dispose()
    '        TBL5.Dispose()
    '    End Try
    'End Function

    Public Function Simpan_Multi(ByVal vdgvItem As DataGridView, ByVal vDgvExchangeRate As Windows.Forms.DataGridView, _
               ByVal vDGVDP As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vDGVPR As DataGridView, _
               ByVal vDGVCostItem As DataGridView, ByVal vFrm As FrmTransaksiAOP_Quick2, _
               ByVal vTblCustom As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1DataTable) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemTableAdapter
        Dim DTA1 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item_exchangerateTableAdapter
        Dim DTA2 As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemdp1TableAdapter
        Dim DTA3 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item_purchasereturnTableAdapter
        Dim DTA4 As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemcostTableAdapter
        Dim Expiry As Date

        Dim DTA5 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item1TableAdapter
        Dim TBL5 As New DseTransaksiAOP.tbltrans_aop_item1DataTable
        Dim row5 As DseTransaksiAOP.tbltrans_aop_item1Row
        Dim AdjTimestamp As Date = #1/1/1950#

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA4.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA5.Connection.ConnectionString = DTA.Connection.ConnectionString

            'simpan data item
            For Each row In vdgvItem.Rows
                If IsDate(row.Cells(vFrm.ColExpiry.Index).Value) = True Then
                    Expiry = row.Cells(vFrm.ColExpiry.Index).Value
                Else
                    Expiry = #1/1/1950#
                End If
                If IsDate(row.Cells(vFrm.ColItemAdjustmentTimestamp.Index).Value) = False Then
                    AdjTimestamp = #1/1/1950#
                Else
                    AdjTimestamp = CDate(row.Cells(vFrm.ColItemAdjustmentTimestamp.Index).Value)
                End If

                'simpan ke basis data item
                DTA.Insert_Multi(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryCodeNo.Index).Value, row.Cells(vFrm.ColPrimaryGRN.Index).Value, _
                row.Cells(vFrm.ColPrimaryPO.Index).Value, row.Cells(vFrm.ColSuppDONo.Index).Value, 0, -1, 0, row.Cells(vFrm.ColTotalQty.Index).Value, _
                row.Cells(vFrm.ColPrimaryUnit.Index).Value, -1, -1, row.Cells(vFrm.ColUnitPrice.Index).Value, row.Cells(vFrm.ColDiscount.Index).Value, _
                row.Cells(vFrm.ColNettPrice.Index).Value, row.Cells(vFrm.ColAmount.Index).Value, row.Cells(vFrm.ColATCost.Index).Value, _
                row.Cells(vFrm.ColRemarksItem.Index).Value, row.Cells(vFrm.ColPrimaryLocation.Index).Value, row.Cells(vFrm.ColLotNumber.Index).Value, _
                Expiry, row.Cells(vFrm.ColReferedToPO.Index).Value, row.Cells(vFrm.ColAtAmount.Index).Value, row.Cells(vFrm.ColDiscAmount.Index).Value, _
                row.Cells(vFrm.ColDiscFromPerct.Index).Value, row.Cells(vFrm.ColTotalQtyMulti.Index).Value, row.Cells(vFrm.ColPrimaryUnitsMulti.Index).Value, _
                row.Cells(vFrm.ColRatioReference.Index).Value, row.Cells(vFrm.ColRatioDefault.Index).Value, CDec(row.Cells(vFrm.ColItemAdjustment.Index).Value), _
                row.Cells(vFrm.ColItemAdjustmentUser.Index).Value, row.Cells(vFrm.ColItemAdjustmentReason.Index).Value, AdjTimestamp, _
                row.Cells(vFrm.ColPriceByMultiUnit.Index).Value)
            Next

            'simpan data item exchange rate
            For Each row In vDgvExchangeRate.Rows
                DTA1.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryDokumen.Index).Value, row.Cells(vFrm.ColPrimaryCurrency1.Index).Value, _
                   row.Cells(vFrm.ColExchangeRate.Index).Value, row.Cells(vFrm.ColExcR_Remarks.Index).Value)
            Next

            'simpan item DP
            For Each row In vDGVDP.Rows
                DTA2.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryAOPDP.Index).Value, row.Cells(vFrm.ColDP_Remarks.Index).Value, 0)
            Next

            'simpan item sales return
            For Each row In vDGVPR.Rows
                DTA3.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryPR.Index).Value, row.Cells(vFrm.ColRemarks_PurchaseReturn.Index).Value)
            Next

            'simpan item cost item
            For Each row In vDGVCostItem.Rows
                If row.Index <> vDGVCostItem.Rows.Count - 1 Then
                    DTA4.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColDescCost.Index).Value, row.Cells(vFrm.ColPrimaryCurrencyCost.Index).Value, row.Cells(vFrm.ColExcRateCost.Index).Value, _
                    row.Cells(vFrm.ColCostNominal.Index).Value, row.Cells(vFrm.ColRemarksCost.Index).Value, row.Cells(vFrm.ColPrimaryAccount.Index).Value)
                End If
            Next

            'simpan primary aop pada tbl grn
            DTA5.Fill(TBL5, vPrimaryTransaksi)
            For Each row5 In TBL5.Rows
                DTA5.UpdateQuery_PrimaryAOP(row5.Primary_AOP, row5.Primary_GRN)
            Next

            Dim ObjCustom As New CustomSMESS.ClsTransaksiAOP_Quick2_AddByCustomProduct
            ObjCustom.SimpanFromTblDummy(vTblCustom, vPrimaryTransaksi)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            DTA2.Dispose()
            DTA4.Dispose()
            DTA3.Dispose()
            DTA5.Dispose()
            TBL5.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemTableAdapter
        Dim DTA1 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item_exchangerateTableAdapter
        Dim DTA2 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item_purchasereturnTableAdapter
        Dim DTA3 As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemdp1TableAdapter
        Dim DTA4 As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemcostTableAdapter
        Dim DTA5 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA4.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA5.Connection.ConnectionString = DTA.Connection.ConnectionString

            'hapus data item
            DTA.Delete(vPrimaryTransaksi)
            DTA1.Delete(vPrimaryTransaksi)
            DTA2.Delete(vPrimaryTransaksi)
            DTA3.Delete(vPrimaryTransaksi)
            DTA4.Delete(vPrimaryTransaksi)
            'update primary aop pada tabel grn
            DTA5.UpdateQuery_SetPrimaryMin(vPrimaryTransaksi)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
            DTA4.Dispose()
            DTA5.Dispose()
        End Try
    End Function

    ' Public Function IsiDGV_Default(ByVal vDGVItem As DataGridView, ByVal vDGVExchangeRate As DataGridView, _
    'ByVal vDGVDP As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vCalculatedBy As _
    'ClsTransaksiAOP.EnumCalculatesBy, ByVal vDGVPR As DataGridView, ByVal vDGVCostItem As DataGridView, ByVal vFrm As FrmTransaksiAOP_Quick2) As Boolean
    '     Dim DTA As New DseTransaksiAOPTableAdapters.tbltransaksi_untukdgv_aopitemTableAdapter
    '     Dim DTA1 As New DseTransaksiAOPTableAdapters.DataTable2TableAdapter
    '     Dim DTA2 As New DseTransaksiAOPTableAdapters.tbltrans_aopitem_purchasereturn_view0TableAdapter
    '     Dim DTA3 As New DseTransaksiAOPTableAdapters.DataTable8TableAdapter

    '     Dim TBL As New DseTransaksiAOP.tbltransaksi_untukdgv_aopitemDataTable
    '     Dim TBL1 As New DseTransaksiAOP.DataTable2DataTable
    '     Dim TBL2 As New DseTransaksiAOP.tbltrans_aopitem_purchasereturn_view0DataTable
    '     Dim TBL3 As New DseTransaksiAOP.DataTable8DataTable

    '     Dim row As DseTransaksiAOP.tbltransaksi_untukdgv_aopitemRow
    '     Dim row1 As DseTransaksiAOP.DataTable2Row
    '     Dim row2 As DseTransaksiAOP.tbltrans_aopitem_purchasereturn_view0Row
    '     Dim row3 As DseTransaksiAOP.DataTable8Row

    '     Dim ObjCD As New Core_Module.ClsCekData
    '     Dim vPrimaryCurrencyRujukan As Double = -1
    '     Dim CurrencyCOdeRujukan As String = ""
    '     Dim rowDGV As DataGridViewRow
    '     Dim DefPrecBase As Integer = 0
    '     Dim ObjCur As New Core_Module.clsMasterCurrency

    '     Try
    '         DTA.Connection.ConnectionString = Core_Module.GetConstring
    '         DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
    '         DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
    '         DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString

    '         DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
    '         'kosongkan dgv
    '         vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '         vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
    '         vDGVItem.Rows.Clear()
    '         vDGVExchangeRate.Rows.Clear()
    '         vDGVDP.Rows.Clear()
    '         vDGVCostItem.Rows.Clear()
    '         vDGVPR.Rows.Clear()

    '         'isikan dgv Cost Item
    '         DTA3.Fill(TBL3, vPrimaryTransaksi)
    '         For Each row3 In TBL3.Rows
    '             vDGVCostItem.Rows.Add("", row3.Cost_Description, row3.Currency_Code, ObjCD.DigitGrouping(row3.Cost, row3.DefPrec2), _
    '             ObjCD.DigitGrouping(row3.Kurs, Core_Module.DefPrecExchangeRate), ObjCD.DigitGrouping(Math.Round(row3.Kurs * row3.Cost, _
    '             DefPrecBase, MidpointRounding.AwayFromZero)), row3.Remarks, row3.Primary_Currency, _
    '             row3.Account_Code, row3.Account_Desc, row3.Primary_Akun)
    '         Next

    '         'isikan tabel item
    '         If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
    '             For Each row In TBL.Rows
    '                 vDGVItem.Rows.Add()
    '                 With row
    '                     rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
    '                     rowDGV.Cells(vFrm.ColNo.Index).Value = (rowDGV.Index + 1).ToString
    '                     rowDGV.Cells(vFrm.ColCodeNo.Index).Value = .Product_Code
    '                     rowDGV.Cells(vFrm.ColDescription.Index).Value = .Description
    '                     rowDGV.Cells(vFrm.ColSuppliersCodeNumber.Index).Value = ""
    '                     rowDGV.Cells(vFrm.ColSuppliersProductDesc.Index).Value = ""
    '                     rowDGV.Cells(vFrm.ColPONO.Index).Value = .PO_No
    '                     rowDGV.Cells(vFrm.ColGRNNO.Index).Value = .GRN_No
    '                     rowDGV.Cells(vFrm.ColStockOutNo.Index).Value = .Stock_In_Purchase
    '                     rowDGV.Cells(vFrm.ColSuppDONo.Index).Value = .Supplier_Do_No
    '                     rowDGV.Cells(vFrm.ColQtyMaster.Index).Value = .Qty_Master
    '                     rowDGV.Cells(vFrm.ColQtyPerMaster.Index).Value = .Qty_Per_Master
    '                     rowDGV.Cells(vFrm.ColMasterUnits.Index).Value = ""
    '                     rowDGV.Cells(vFrm.ColQty.Index).Value = .Qty
    '                     rowDGV.Cells(vFrm.ColTotalQty.Index).Value = ObjCD.DigitGrouping(.Total_Qty, .Decimal_Product)
    '                     rowDGV.Cells(vFrm.ColUnit.Index).Value = .Units
    '                     rowDGV.Cells(vFrm.ColSecQty.Index).Value = -1
    '                     rowDGV.Cells(vFrm.ColSecUnit.Index).Value = ""
    '                     rowDGV.Cells(vFrm.ColCurrency.Index).Value = .Currency_Code
    '                     rowDGV.Cells(vFrm.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(.Unit_Price, .Decimal_Currency)
    '                     rowDGV.Cells(vFrm.ColDiscount.Index).Value = ObjCD.DigitGrouping(.Discount, 2)
    '                     rowDGV.Cells(vFrm.ColDiscAmount.Index).Value = ObjCD.DigitGrouping(.Disc_Amount, .Decimal_Currency)
    '                     rowDGV.Cells(vFrm.ColNettPrice.Index).Value = ObjCD.DigitGrouping(.Nett_Price, .Decimal_Currency)
    '                     rowDGV.Cells(vFrm.ColAmount.Index).Value = ObjCD.DigitGrouping(.Amount, .Decimal_Currency)
    '                     rowDGV.Cells(vFrm.ColATCost.Index).Value = ObjCD.DigitGrouping(.At_Price, 0)
    '                     rowDGV.Cells(vFrm.ColShowDetail.Index).Value = "Show"
    '                     rowDGV.Cells(vFrm.ColRemarksItem.Index).Value = .Remarks
    '                     rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value = .Primary_Product
    '                     rowDGV.Cells(vFrm.ColPrimaryPO.Index).Value = .Primary_PO
    '                     rowDGV.Cells(vFrm.ColPrimaryGRN.Index).Value = .Primary_GRN
    '                     rowDGV.Cells(vFrm.ColPrimaryMasterUnits.Index).Value = -1
    '                     rowDGV.Cells(vFrm.ColPrimaryUnit.Index).Value = .Primary_Units
    '                     rowDGV.Cells(vFrm.ColPrimarySecUnit.Index).Value = .Primary_Secondary_Units
    '                     rowDGV.Cells(vFrm.ColPrimaryCurrency.Index).Value = .Primary_Currency
    '                     If .Storage_Area_Code = "" Then
    '                         .Storage_Area_Code = "-"
    '                     End If
    '                     rowDGV.Cells(vFrm.ColLocation.Index).Value = .Storage_Area_Code
    '                     rowDGV.Cells(vFrm.ColPrimaryLocation.Index).Value = .Primary_Location
    '                     rowDGV.Cells(vFrm.ColLotNumber.Index).Value = .BatchNumber
    '                     If .Expiry_Date = #1/1/1950# Then
    '                         rowDGV.Cells(vFrm.ColExpiry.Index).Value = "-"
    '                     Else
    '                         rowDGV.Cells(vFrm.ColExpiry.Index).Value = Format(.Expiry_Date, "dd MMMM yyyy")
    '                     End If
    '                     rowDGV.Cells(vFrm.ColReferedToPO.Index).Value = .Refered_To_PO
    '                     If rowDGV.Cells(vFrm.ColReferedToPO.Index).Value = 0 Then
    '                         rowDGV.Cells(vFrm.ColPONO.Index).Value = ""
    '                     End If
    '                     rowDGV.Cells(vFrm.ColAtAmount.Index).Value = .At_Amount
    '                     rowDGV.Cells(vFrm.ColDiscFromPerct.Index).Value = .Disc_From_Perct
    '                 End With
    '             Next

    '             row = TBL.Rows(0)
    '             vPrimaryCurrencyRujukan = row.Primary_Currency
    '             CurrencyCOdeRujukan = row.Currency_Code
    '         End If
    '         'vDGVItem.Columns(vFrm.ColPONO.Index).Visible = False
    '         'vDGVItem.Columns(vFrm.ColGRNNO.Index).Visible = False
    '         'vDGVItem.Columns(vFrm.ColStockOutNo.Index).Visible = False

    '         'vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
    '         vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

    '         'isikan tabel item dp
    '         If DTA1.FillBy_PrimaryAOP(TBL1, vPrimaryTransaksi) > 0 Then
    '             For Each row1 In TBL1.Rows
    '                 With row1
    '                     vDGVDP.Rows.Add("0", .No_AOP_DP, .Currency_Code, .Sub_Total, .Remarks, _
    '                     .Primary_AOP_DP)
    '                 End With
    '             Next
    '         End If

    '         'isikan tabel item exchange rate
    '         If vCalculatedBy = ClsTransaksiAOP.EnumCalculatesBy.AOP Then
    '             Dim DTAEX As New DseTransaksiAOPTableAdapters.DataTable3TableAdapter
    '             Dim TBLEX As New DseTransaksiAOP.DataTable3DataTable
    '             Dim RowEX As DseTransaksiAOP.DataTable3Row
    '             If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
    '                 For Each RowEX In TBLEX.Rows
    '                     With RowEX
    '                         vDGVExchangeRate.Rows.Add("0", .NoTransaksiAOP, CurrencyCOdeRujukan, _
    '                         .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, vPrimaryCurrencyRujukan)
    '                     End With
    '                 Next
    '             End If
    '             DTAEX.Dispose()
    '             TBLEX.Dispose()
    '         ElseIf vCalculatedBy = ClsTransaksiAOP.EnumCalculatesBy.GRN Then
    '             Dim DTAEX As New DseTransaksiAOPTableAdapters.DataTable4TableAdapter
    '             Dim TBLEX As New DseTransaksiAOP.DataTable4DataTable
    '             Dim RowEX As DseTransaksiAOP.DataTable4Row
    '             If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
    '                 For Each RowEX In TBLEX.Rows
    '                     With RowEX
    '                         vDGVExchangeRate.Rows.Add("0", .GRN_No, CurrencyCOdeRujukan, _
    '                         .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, vPrimaryCurrencyRujukan)
    '                     End With
    '                 Next
    '             End If
    '             DTAEX.Dispose()
    '             TBLEX.Dispose()
    '         ElseIf vCalculatedBy = ClsTransaksiAOP.EnumCalculatesBy.PO Then
    '             Dim DTAEX As New DseTransaksiAOPTableAdapters.DataTable5TableAdapter
    '             Dim TBLEX As New DseTransaksiAOP.DataTable5DataTable
    '             Dim RowEX As DseTransaksiAOP.DataTable5Row
    '             If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
    '                 For Each RowEX In TBLEX.Rows
    '                     With RowEX
    '                         vDGVExchangeRate.Rows.Add("0", .PO_No, CurrencyCOdeRujukan, _
    '                         .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, vPrimaryCurrencyRujukan)
    '                     End With
    '                 Next
    '             End If
    '             DTAEX.Dispose()
    '             TBLEX.Dispose()
    '         End If

    '         'isikan dgv PR
    '         DTA2.FillBy_PrimaryAOP(TBL2, vPrimaryTransaksi)
    '         For Each row2 In TBL2.Rows
    '             vDGVPR.Rows.Add("", row2.PruchaseReturnNo, row2.GRN_No, row2.Remarks, row2.Primary_PR)
    '         Next

    '         'set nomer
    '         ObjCD.SetNumber(vDGVDP)
    '         ObjCD.SetNumber(vDGVExchangeRate)
    '         ObjCD.SetNumber(vDGVPR)
    '         ObjCD.SetNumber(vDGVCostItem)

    '         'nilai balikan
    '         Return True
    '     Catch ex As Exception
    '         MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '     Finally
    '         DTA.Dispose()
    '         TBL = Nothing
    '         ObjCD = Nothing
    '         ObjCur = Nothing
    '         DTA1.Dispose()
    '         DTA2.Dispose()
    '         DTA3.Dispose()
    '     End Try
    ' End Function

    ''@rosadinnn
    ''@MUaT
    Public Function IsiDGV_Multi(ByVal vDGVItem As DataGridView, ByVal vDGVExchangeRate As DataGridView, _
   ByVal vDGVDP As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vCalculatedBy As _
   ClsTransaksiAOP.EnumCalculatesBy, ByVal vDGVPR As DataGridView, ByVal vDGVCostItem As DataGridView, ByVal vFrm As FrmTransaksiAOP_Quick2) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltransaksi_untukdgv_aopitem_MultiTableAdapter
        Dim DTA1 As New DseTransaksiAOPTableAdapters.DataTable2TableAdapter
        Dim DTA2 As New DseTransaksiAOPTableAdapters.tbltrans_aopitem_purchasereturn_view0TableAdapter
        Dim DTA3 As New DseTransaksiAOPTableAdapters.DataTable8TableAdapter

        Dim TBL As New DseTransaksiAOP.tbltransaksi_untukdgv_aopitem_MultiDataTable
        Dim TBL1 As New DseTransaksiAOP.DataTable2DataTable
        Dim TBL2 As New DseTransaksiAOP.tbltrans_aopitem_purchasereturn_view0DataTable
        Dim TBL3 As New DseTransaksiAOP.DataTable8DataTable

        Dim row As DseTransaksiAOP.tbltransaksi_untukdgv_aopitem_MultiRow
        Dim row1 As DseTransaksiAOP.DataTable2Row
        Dim row2 As DseTransaksiAOP.tbltrans_aopitem_purchasereturn_view0Row
        Dim row3 As DseTransaksiAOP.DataTable8Row

        Dim ObjCD As New Core_Module.ClsCekData
        Dim vPrimaryCurrencyRujukan As Double = -1
        Dim CurrencyCOdeRujukan As String = ""
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim objProduct As New Core_Module.clsMasterProduct

        Dim TBLMultiUnit As New DseTransaksiAOP_CreatByStockInPurchase.tblmaster_product_item_multiunitDataTable

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString

            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItem.Rows.Clear()
            vDGVExchangeRate.Rows.Clear()
            vDGVDP.Rows.Clear()
            vDGVCostItem.Rows.Clear()
            vDGVPR.Rows.Clear()

            'isikan dgv Cost Item
            DTA3.Fill(TBL3, vPrimaryTransaksi)
            For Each row3 In TBL3.Rows
                vDGVCostItem.Rows.Add("", row3.Cost_Description, row3.Currency_Code, ObjCD.DigitGrouping(row3.Cost, row3.DefPrec2), _
                ObjCD.DigitGrouping(row3.Kurs, Core_Module.DefPrecExchangeRate), ObjCD.DigitGrouping(Math.Round(row3.Kurs * row3.Cost, _
                DefPrecBase, MidpointRounding.AwayFromZero)), row3.Remarks, row3.Primary_Currency, _
                row3.Account_Code, row3.Account_Desc, row3.Primary_Akun)
            Next

            'isikan tabel item
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                Me.IsikanDecimalMulti(TBL, TBLMultiUnit)

                For Each row In TBL.Rows
                    vDGVItem.Rows.Add()
                    With row
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(vFrm.ColNo.Index).Value = (rowDGV.Index + 1).ToString
                        rowDGV.Cells(vFrm.ColCodeNo.Index).Value = .Product_Code
                        rowDGV.Cells(vFrm.ColDescription.Index).Value = .Description
                        rowDGV.Cells(vFrm.ColSuppliersCodeNumber.Index).Value = ""
                        rowDGV.Cells(vFrm.ColSuppliersProductDesc.Index).Value = ""
                        rowDGV.Cells(vFrm.ColPONO.Index).Value = .PO_No
                        rowDGV.Cells(vFrm.ColGRNNO.Index).Value = .GRN_No
                        rowDGV.Cells(vFrm.ColStockOutNo.Index).Value = .Stock_In_Purchase
                        rowDGV.Cells(vFrm.ColSuppDONo.Index).Value = .Supplier_Do_No
                        rowDGV.Cells(vFrm.ColQtyMaster.Index).Value = .Qty_Master
                        rowDGV.Cells(vFrm.ColQtyPerMaster.Index).Value = .Qty_Per_Master
                        rowDGV.Cells(vFrm.ColMasterUnits.Index).Value = ""
                        rowDGV.Cells(vFrm.ColQty.Index).Value = .Qty
                        rowDGV.Cells(vFrm.ColTotalQty.Index).Value = ObjCD.DigitGrouping(.Total_Qty, .Decimal_Product)
                        rowDGV.Cells(vFrm.ColUnit.Index).Value = .Units
                        rowDGV.Cells(vFrm.ColSecQty.Index).Value = -1
                        rowDGV.Cells(vFrm.ColSecUnit.Index).Value = ""
                        rowDGV.Cells(vFrm.ColCurrency.Index).Value = .Currency_Code
                        rowDGV.Cells(vFrm.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(.Unit_Price, .Decimal_Currency)
                        rowDGV.Cells(vFrm.ColDiscount.Index).Value = ObjCD.DigitGrouping(.Discount, 2)
                        rowDGV.Cells(vFrm.ColDiscAmount.Index).Value = ObjCD.DigitGrouping(.Disc_Amount, .Decimal_Currency)
                        rowDGV.Cells(vFrm.ColNettPrice.Index).Value = ObjCD.DigitGrouping(.Nett_Price, .Decimal_Currency)
                        rowDGV.Cells(vFrm.ColAmount.Index).Value = ObjCD.DigitGrouping(.Amount, .Decimal_Currency)
                        rowDGV.Cells(vFrm.ColATCost.Index).Value = ObjCD.DigitGrouping(.At_Price, 0)
                        rowDGV.Cells(vFrm.ColShowDetail.Index).Value = "Show"
                        rowDGV.Cells(vFrm.ColRemarksItem.Index).Value = .Remarks
                        rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rowDGV.Cells(vFrm.ColPrimaryPO.Index).Value = .Primary_PO
                        rowDGV.Cells(vFrm.ColPrimaryGRN.Index).Value = .Primary_GRN
                        rowDGV.Cells(vFrm.ColPrimaryMasterUnits.Index).Value = -1
                        rowDGV.Cells(vFrm.ColPrimaryUnit.Index).Value = .Primary_Units
                        rowDGV.Cells(vFrm.ColPrimarySecUnit.Index).Value = .Primary_Secondary_Units
                        rowDGV.Cells(vFrm.ColPrimaryCurrency.Index).Value = .Primary_Currency
                        If .Storage_Area_Code = "" Then
                            .Storage_Area_Code = "-"
                        End If
                        rowDGV.Cells(vFrm.ColLocation.Index).Value = .Storage_Area_Code
                        rowDGV.Cells(vFrm.ColPrimaryLocation.Index).Value = .Primary_Location
                        rowDGV.Cells(vFrm.ColLotNumber.Index).Value = .BatchNumber
                        If .Expiry_Date = #1/1/1950# Then
                            rowDGV.Cells(vFrm.ColExpiry.Index).Value = "-"
                        Else
                            rowDGV.Cells(vFrm.ColExpiry.Index).Value = Format(.Expiry_Date, "dd MMMM yyyy")
                        End If
                        rowDGV.Cells(vFrm.ColReferedToPO.Index).Value = .Refered_To_PO
                        If rowDGV.Cells(vFrm.ColReferedToPO.Index).Value = 0 Then
                            rowDGV.Cells(vFrm.ColPONO.Index).Value = ""
                        End If
                        rowDGV.Cells(vFrm.ColAtAmount.Index).Value = .At_Amount
                        rowDGV.Cells(vFrm.ColDiscFromPerct.Index).Value = .Disc_From_Perct
                        If .Primary_Units_Multi = .Primary_Units Then
                            rowDGV.Cells(vFrm.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, .Decimal_Product)
                        Else
                            rowDGV.Cells(vFrm.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, Me.GetDecimalMulti(TBLMultiUnit, .Primary_Product, .Primary_Units_Multi))
                        End If
                        rowDGV.Cells(vFrm.ColUnitsCodeMulti.Index).Value = .Unit_Code_Multi
                        rowDGV.Cells(vFrm.ColRatioReference.Index).Value = .Ratio_Reference
                        rowDGV.Cells(vFrm.ColRatioDefault.Index).Value = .Ratio_Default
                        rowDGV.Cells(vFrm.ColPrimaryUnitsMulti.Index).Value = .Primary_Units_Multi
                        rowDGV.Cells(vFrm.ColItemAdjustment.Index).Value = ObjCD.DigitGrouping(.Adjustment, .Decimal_Currency)
                        rowDGV.Cells(vFrm.ColItemAdjustmentReason.Index).Value = .Reason
                        rowDGV.Cells(vFrm.ColItemAdjustmentUser.Index).Value = .User_Code
                        If .Adjustment_Timestamp = #1/1/1950# Then
                            rowDGV.Cells(vFrm.ColItemAdjustmentTimestamp.Index).Value = "-"
                        Else
                            rowDGV.Cells(vFrm.ColItemAdjustmentTimestamp.Index).Value = Format(.Adjustment_Timestamp, "dd MMM yyyy HH:mm")
                        End If

                        If .Primary_Units <> .Primary_Units_Multi Then
                            rowDGV.Cells(vFrm.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, _
                                    Me.GetDecimalMulti(TBLMultiUnit, .Primary_Product, .Primary_Units_Multi)) _
                                    & " " & .Unit_Code_Multi & " = " & ObjCD.DigitGrouping(.Ratio_Default, _
                                    .Decimal_Product) & " " & .Units
                        Else
                            rowDGV.Cells(vFrm.ColRatioDesc.Index).Value = "-"
                        End If
                        rowDGV.Cells(vFrm.ColPriceByMultiUnit.Index).Value = .price_by_multi_unit
                        If .price_by_multi_unit = 1 Then
                            rowDGV.Cells(vFrm.ColPriceByMultiUnit_Str.Index).Value = "/" + .Unit_Code_Multi
                        Else
                            rowDGV.Cells(vFrm.ColPriceByMultiUnit_Str.Index).Value = "/" + .Units
                        End If
                    End With
                Next

                row = TBL.Rows(0)
                vPrimaryCurrencyRujukan = row.Primary_Currency
                CurrencyCOdeRujukan = row.Currency_Code
            End If
            'vDGVItem.Columns(vFrm.ColPONO.Index).Visible = False
            'vDGVItem.Columns(vFrm.ColGRNNO.Index).Visible = False
            'vDGVItem.Columns(vFrm.ColStockOutNo.Index).Visible = False

            'vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'isikan tabel item dp
            If DTA1.FillBy_PrimaryAOP(TBL1, vPrimaryTransaksi) > 0 Then
                For Each row1 In TBL1.Rows
                    With row1
                        vDGVDP.Rows.Add("0", .No_AOP_DP, .Currency_Code, .Sub_Total, .Remarks, _
                        .Primary_AOP_DP)
                    End With
                Next
            End If

            'isikan tabel item exchange rate
            If vCalculatedBy = ClsTransaksiAOP.EnumCalculatesBy.AOP Then
                Dim DTAEX As New DseTransaksiAOPTableAdapters.DataTable3TableAdapter
                Dim TBLEX As New DseTransaksiAOP.DataTable3DataTable
                Dim RowEX As DseTransaksiAOP.DataTable3Row
                If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
                    For Each RowEX In TBLEX.Rows
                        With RowEX
                            vDGVExchangeRate.Rows.Add("0", .NoTransaksiAOP, CurrencyCOdeRujukan, _
                            .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, vPrimaryCurrencyRujukan)
                        End With
                    Next
                End If
                DTAEX.Dispose()
                TBLEX.Dispose()
            ElseIf vCalculatedBy = ClsTransaksiAOP.EnumCalculatesBy.GRN Then
                Dim DTAEX As New DseTransaksiAOPTableAdapters.DataTable4TableAdapter
                Dim TBLEX As New DseTransaksiAOP.DataTable4DataTable
                Dim RowEX As DseTransaksiAOP.DataTable4Row
                If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
                    For Each RowEX In TBLEX.Rows
                        With RowEX
                            vDGVExchangeRate.Rows.Add("0", .GRN_No, CurrencyCOdeRujukan, _
                            .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, vPrimaryCurrencyRujukan)
                        End With
                    Next
                End If
                DTAEX.Dispose()
                TBLEX.Dispose()
            ElseIf vCalculatedBy = ClsTransaksiAOP.EnumCalculatesBy.PO Then
                Dim DTAEX As New DseTransaksiAOPTableAdapters.DataTable5TableAdapter
                Dim TBLEX As New DseTransaksiAOP.DataTable5DataTable
                Dim RowEX As DseTransaksiAOP.DataTable5Row
                If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
                    For Each RowEX In TBLEX.Rows
                        With RowEX
                            vDGVExchangeRate.Rows.Add("0", .PO_No, CurrencyCOdeRujukan, _
                            .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, vPrimaryCurrencyRujukan)
                        End With
                    Next
                End If
                DTAEX.Dispose()
                TBLEX.Dispose()
            End If

            'isikan dgv PR
            DTA2.FillBy_PrimaryAOP(TBL2, vPrimaryTransaksi)
            For Each row2 In TBL2.Rows
                vDGVPR.Rows.Add("", row2.PruchaseReturnNo, row2.GRN_No, row2.Remarks, row2.Primary_PR)
            Next

            'set nomer
            ObjCD.SetNumber(vDGVDP)
            ObjCD.SetNumber(vDGVExchangeRate)
            ObjCD.SetNumber(vDGVPR)
            ObjCD.SetNumber(vDGVCostItem)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
            DTA1.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
        End Try
    End Function

    'Public Function IsiDGV_Default(ByVal vDGVItem As DataGridView, ByVal vDGVExchangeRate As DataGridView, _
    '      ByVal vDGVDP As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vCalculatedBy As _
    '      ClsTransaksiAOP.EnumCalculatesBy, ByVal vDGVPR As DataGridView, ByVal vDGVCostItem As DataGridView, ByVal vFrm As FrmTransaksiAOP2) As Boolean
    '    Dim DTA As New DseTransaksiAOPTableAdapters.tbltransaksi_untukdgv_aopitemTableAdapter
    '    Dim DTA1 As New DseTransaksiAOPTableAdapters.DataTable2TableAdapter
    '    Dim DTA2 As New DseTransaksiAOPTableAdapters.tbltrans_aopitem_purchasereturn_view0TableAdapter
    '    Dim DTA3 As New DseTransaksiAOPTableAdapters.DataTable8TableAdapter

    '    Dim TBL As New DseTransaksiAOP.tbltransaksi_untukdgv_aopitemDataTable
    '    Dim TBL1 As New DseTransaksiAOP.DataTable2DataTable
    '    Dim TBL2 As New DseTransaksiAOP.tbltrans_aopitem_purchasereturn_view0DataTable
    '    Dim TBL3 As New DseTransaksiAOP.DataTable8DataTable

    '    Dim row As DseTransaksiAOP.tbltransaksi_untukdgv_aopitemRow
    '    Dim row1 As DseTransaksiAOP.DataTable2Row
    '    Dim row2 As DseTransaksiAOP.tbltrans_aopitem_purchasereturn_view0Row
    '    Dim row3 As DseTransaksiAOP.DataTable8Row

    '    Dim ObjCD As New Core_Module.ClsCekData
    '    Dim ObjCur As New Core_Module.clsMasterCurrency

    '    Dim vPrimaryCurrencyRujukan As Double = -1
    '    Dim CurrencyCOdeRujukan As String = ""
    '    Dim DefPrecBase As Integer = 0

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
    '        DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
    '        DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString

    '        DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
    '        'kosongkan dgv
    '        vDGVItem.Rows.Clear()
    '        vDGVExchangeRate.Rows.Clear()
    '        vDGVDP.Rows.Clear()
    '        vDGVCostItem.Rows.Clear()
    '        vDGVPR.Rows.Clear()

    '        'isikan dgv Cost Item
    '        DTA3.Fill(TBL3, vPrimaryTransaksi)
    '        For Each row3 In TBL3.Rows
    '            vDGVCostItem.Rows.Add("", row3.Cost_Description, row3.Currency_Code, ObjCD.DigitGrouping(row3.Cost, row3.DefPrec2), _
    '            ObjCD.DigitGrouping(row3.Kurs, Core_Module.DefPrecExchangeRate), ObjCD.DigitGrouping(Math.Round(row3.Kurs * row3.Cost, _
    '            DefPrecBase, MidpointRounding.AwayFromZero)), row3.Remarks, row3.Primary_Currency, _
    '            row3.Account_Code, row3.Account_Desc, row3.Primary_Akun)
    '        Next

    '        'isikan tabel item
    '        vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
    '        vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '        If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
    '            For Each row In TBL.Rows
    '                With row
    '                    vDGVItem.Rows.Add((vDGVItem.Rows.Count + 1).ToString, .Product_Code, .Description, ._Supplier_sProductCode_, ._Supplier_sProductDesc_, _
    '                    .PO_No, .GRN_No, .Stock_In_Purchase, .Supplier_Do_No, .Qty_Master, .Qty_Per_Master, .Master_Unit, .Qty, _
    '                    ObjCD.DigitGrouping(.Total_Qty, .Decimal_Product), .Units, .BatchNumber, .Secondary_Qty, .Secondary_Unit, .Currency_Code, ObjCD.DigitGrouping(.Unit_Price, .Decimal_Currency), _
    '                    ObjCD.DigitGrouping(.Discount, 2), ObjCD.DigitGrouping(.Disc_Amount, .Decimal_Currency), ObjCD.DigitGrouping(.Nett_Price, .Decimal_Currency), _
    '                    ObjCD.DigitGrouping(.Amount, .Decimal_Currency), ObjCD.DigitGrouping(.At_Price, DefPrecBase), "Show", .Remarks, _
    '                    .Primary_Product, .Primary_PO, .Primary_GRN, .Primary_Master_Unit, .Primary_Units, _
    '                    .Primary_Secondary_Units, .Primary_Currency, ObjCD.DigitGrouping(.At_Amount, DefPrecBase), .Disc_From_Perct)
    '                End With
    '            Next

    '            row = TBL.Rows(0)
    '            vPrimaryCurrencyRujukan = row.Primary_Currency
    '            CurrencyCOdeRujukan = row.Currency_Code
    '        End If
    '        vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

    '        'isikan tabel item dp
    '        If DTA1.FillBy_PrimaryAOP(TBL1, vPrimaryTransaksi) > 0 Then
    '            For Each row1 In TBL1.Rows
    '                With row1
    '                    vDGVDP.Rows.Add("0", .No_AOP_DP, .Currency_Code, ObjCD.DigitGrouping(.Sub_Total, .Default_Precision), .Remarks, _
    '                    .Primary_AOP_DP)
    '                End With
    '            Next
    '        End If

    '        'isikan tabel item exchange rate
    '        If vCalculatedBy = ClsTransaksiAOP.EnumCalculatesBy.AOP Then
    '            Dim DTAEX As New DseTransaksiAOPTableAdapters.DataTable3TableAdapter
    '            Dim TBLEX As New DseTransaksiAOP.DataTable3DataTable
    '            Dim RowEX As DseTransaksiAOP.DataTable3Row
    '            If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
    '                For Each RowEX In TBLEX.Rows
    '                    With RowEX
    '                        vDGVExchangeRate.Rows.Add("0", .NoTransaksiAOP, CurrencyCOdeRujukan, _
    '                        .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, vPrimaryCurrencyRujukan)
    '                    End With
    '                Next
    '            End If
    '            DTAEX.Dispose()
    '            TBLEX.Dispose()
    '        ElseIf vCalculatedBy = ClsTransaksiAOP.EnumCalculatesBy.GRN Then
    '            Dim DTAEX As New DseTransaksiAOPTableAdapters.DataTable4TableAdapter
    '            Dim TBLEX As New DseTransaksiAOP.DataTable4DataTable
    '            Dim RowEX As DseTransaksiAOP.DataTable4Row
    '            If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
    '                For Each RowEX In TBLEX.Rows
    '                    With RowEX
    '                        vDGVExchangeRate.Rows.Add("0", .GRN_No, CurrencyCOdeRujukan, _
    '                        .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, vPrimaryCurrencyRujukan)
    '                    End With
    '                Next
    '            End If
    '            DTAEX.Dispose()
    '            TBLEX.Dispose()
    '        ElseIf vCalculatedBy = ClsTransaksiAOP.EnumCalculatesBy.PO Then
    '            Dim DTAEX As New DseTransaksiAOPTableAdapters.DataTable5TableAdapter
    '            Dim TBLEX As New DseTransaksiAOP.DataTable5DataTable
    '            Dim RowEX As DseTransaksiAOP.DataTable5Row
    '            If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
    '                For Each RowEX In TBLEX.Rows
    '                    With RowEX
    '                        vDGVExchangeRate.Rows.Add("0", .PO_No, CurrencyCOdeRujukan, _
    '                        .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, vPrimaryCurrencyRujukan)
    '                    End With
    '                Next
    '            End If
    '            DTAEX.Dispose()
    '            TBLEX.Dispose()
    '        End If

    '        'isikan dgv PR
    '        DTA2.FillBy_PrimaryAOP(TBL2, vPrimaryTransaksi)
    '        For Each row2 In TBL2.Rows
    '            vDGVPR.Rows.Add("", row2.PruchaseReturnNo, row2.GRN_No, row2.Remarks, row2.Primary_PR)
    '        Next

    '        'set nomer
    '        ObjCD.SetNumber(vDGVDP)
    '        ObjCD.SetNumber(vDGVExchangeRate)
    '        ObjCD.SetNumber(vDGVPR)
    '        ObjCD.SetNumber(vDGVCostItem)

    '        'nilai balikan
    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '        TBL = Nothing
    '        ObjCD = Nothing
    '        ObjCur = Nothing
    '        DTA1.Dispose()
    '        DTA2.Dispose()
    '        DTA3.Dispose()
    '    End Try
    'End Function

    ''@rosadinnn
    ''@MUaT
    Private Sub IsikanDecimalMulti(ByVal vTBL As DseTransaksiAOP.tbltransaksi_untukdgv_aopitem_MultiDataTable, _
    ByVal vTBLMultiUnit As DseTransaksiAOP_CreatByStockInPurchase.tblmaster_product_item_multiunitDataTable)
        Dim row As DseTransaksiAOP.tbltransaksi_untukdgv_aopitem_MultiRow
        Dim DTAMulti As New DseTransaksiAOP_CreatByStockInPurchaseTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Dim TBLMulti As New DseTransaksiAOP_CreatByStockInPurchase.tblmaster_product_item_multiunitDataTable
        Dim rowMultiUnit, rowMultiUnit1 As DseTransaksiAOP_CreatByStockInPurchase.tblmaster_product_item_multiunitRow
        Dim ListPrimary As New List(Of Double)

        Try
            'fil list primary
            For Each row In vTBL.Rows
                If ListPrimary.Contains(row.Primary_Product) = False Then
                    ListPrimary.Add(row.Primary_Product)
                End If
            Next

            'fill tabel multi unit
            vTBLMultiUnit.Rows.Clear()
            For Each PrimaryProduct As Double In ListPrimary
                TBLMulti.Rows.Clear()
                DTAMulti.Fill(TBLMulti, PrimaryProduct)
                For Each rowMultiUnit In TBLMulti.Rows
                    rowMultiUnit1 = vTBLMultiUnit.Newtblmaster_product_item_multiunitRow
                    For ij As Integer = 0 To vTBLMultiUnit.Columns.Count - 1
                        rowMultiUnit1.Item(ij) = rowMultiUnit.Item(ij)
                    Next
                    vTBLMultiUnit.Rows.Add(rowMultiUnit1)
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetDecimalMulti(ByVal vTBLMultiUnit As DseTransaksiAOP_CreatByStockInPurchase.tblmaster_product_item_multiunitDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double) As Integer
        Dim rowMultiUnit As DseTransaksiAOP_CreatByStockInPurchase.tblmaster_product_item_multiunitRow
        Try
            For Each rowMultiUnit In vTBLMultiUnit.Rows
                If rowMultiUnit.Primary_Product = vPrimaryProduct And rowMultiUnit.Primary_Unit = vPrimaryUnit Then
                    Return rowMultiUnit.Default_Precision
                End If
            Next

            MsgBox("Logic error: no multi unit detected. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function IsiDGV_Multi(ByVal vDGVItem As DataGridView, ByVal vDGVExchangeRate As DataGridView, _
          ByVal vDGVDP As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vCalculatedBy As _
          ClsTransaksiAOP.EnumCalculatesBy, ByVal vDGVPR As DataGridView, ByVal vDGVCostItem As DataGridView, ByVal vFrm As FrmTransaksiAOP2) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltransaksi_untukdgv_aopitem_MultiTableAdapter
        Dim DTA1 As New DseTransaksiAOPTableAdapters.DataTable2TableAdapter
        Dim DTA2 As New DseTransaksiAOPTableAdapters.tbltrans_aopitem_purchasereturn_view0TableAdapter
        Dim DTA3 As New DseTransaksiAOPTableAdapters.DataTable8TableAdapter

        Dim TBL As New DseTransaksiAOP.tbltransaksi_untukdgv_aopitem_MultiDataTable
        Dim TBL1 As New DseTransaksiAOP.DataTable2DataTable
        Dim TBL2 As New DseTransaksiAOP.tbltrans_aopitem_purchasereturn_view0DataTable
        Dim TBL3 As New DseTransaksiAOP.DataTable8DataTable

        Dim row As DseTransaksiAOP.tbltransaksi_untukdgv_aopitem_MultiRow
        Dim row1 As DseTransaksiAOP.DataTable2Row
        Dim row2 As DseTransaksiAOP.tbltrans_aopitem_purchasereturn_view0Row
        Dim row3 As DseTransaksiAOP.DataTable8Row

        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Dim vPrimaryCurrencyRujukan As Double = -1
        Dim CurrencyCOdeRujukan As String = ""
        Dim DefPrecBase As Integer = 0

        Dim rowDgv As DataGridViewRow
        Dim objProduct As New Core_Module.clsMasterProduct

        Dim TBLMultiUnit As New DseTransaksiAOP_CreatByStockInPurchase.tblmaster_product_item_multiunitDataTable

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString

            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            vDGVExchangeRate.Rows.Clear()
            vDGVDP.Rows.Clear()
            vDGVCostItem.Rows.Clear()
            vDGVPR.Rows.Clear()

            'isikan dgv Cost Item
            DTA3.Fill(TBL3, vPrimaryTransaksi)
            For Each row3 In TBL3.Rows
                vDGVCostItem.Rows.Add("", row3.Cost_Description, row3.Currency_Code, ObjCD.DigitGrouping(row3.Cost, row3.DefPrec2), _
                ObjCD.DigitGrouping(row3.Kurs, Core_Module.DefPrecExchangeRate), ObjCD.DigitGrouping(Math.Round(row3.Kurs * row3.Cost, _
                DefPrecBase, MidpointRounding.AwayFromZero)), row3.Remarks, row3.Primary_Currency, _
                row3.Account_Code, row3.Account_Desc, row3.Primary_Akun)
            Next

            'isikan tabel item
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                Me.IsikanDecimalMulti(TBL, TBLMultiUnit)

                For Each row In TBL.Rows

                    vDGVItem.Rows.Add()
                    rowDgv = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                    With row
                        rowDgv.Cells(vFrm.ColItem_No.Index).Value = (vDGVItem.Rows.Count + 1).ToString
                        rowDgv.Cells(vFrm.ColCodeNo.Index).Value = .Product_Code
                        rowDgv.Cells(vFrm.ColDescription.Index).Value = .Description
                        rowDgv.Cells(vFrm.ColSuppliersCodeNumber.Index).Value = ._Supplier_sProductCode_
                        rowDgv.Cells(vFrm.ColSuppliersProductDesc.Index).Value = ._Supplier_sProductDesc_
                        rowDgv.Cells(vFrm.ColPONO.Index).Value = .PO_No
                        rowDgv.Cells(vFrm.ColGRNNO.Index).Value = .GRN_No
                        rowDgv.Cells(vFrm.ColStockOutNo.Index).Value = .Stock_In_Purchase
                        rowDgv.Cells(vFrm.ColSuppDONo.Index).Value = .Supplier_Do_No
                        rowDgv.Cells(vFrm.ColQtyMaster.Index).Value = .Qty_Master
                        rowDgv.Cells(vFrm.ColQtyPerMaster.Index).Value = .Qty_Per_Master
                        rowDgv.Cells(vFrm.ColMasterUnits.Index).Value = .Master_Unit
                        rowDgv.Cells(vFrm.ColQty.Index).Value = .Qty
                        rowDgv.Cells(vFrm.ColTotalQty.Index).Value = ObjCD.DigitGrouping(.Total_Qty, .Decimal_Product)
                        rowDgv.Cells(vFrm.ColUnit.Index).Value = .Units
                        rowDgv.Cells(vFrm.ColBatchNo.Index).Value = .BatchNumber
                        rowDgv.Cells(vFrm.ColSecQty.Index).Value = .Secondary_Qty
                        rowDgv.Cells(vFrm.ColSecUnit.Index).Value = .Secondary_Unit
                        rowDgv.Cells(vFrm.ColCurrency.Index).Value = .Currency_Code
                        rowDgv.Cells(vFrm.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(.Unit_Price, .Decimal_Currency)
                        rowDgv.Cells(vFrm.ColDiscount.Index).Value = ObjCD.DigitGrouping(.Discount, 2)
                        rowDgv.Cells(vFrm.ColDiscAmount.Index).Value = ObjCD.DigitGrouping(.Disc_Amount, .Decimal_Currency)
                        rowDgv.Cells(vFrm.ColNettPrice.Index).Value = ObjCD.DigitGrouping(.Nett_Price, .Decimal_Currency)
                        rowDgv.Cells(vFrm.ColAmount.Index).Value = ObjCD.DigitGrouping(.Amount, .Decimal_Currency)
                        rowDgv.Cells(vFrm.ColATCost.Index).Value = ObjCD.DigitGrouping(.At_Price, DefPrecBase)
                        rowDgv.Cells(vFrm.ColShowDetail.Index).Value = "Show"
                        rowDgv.Cells(vFrm.ColRemarks.Index).Value = .Remarks
                        rowDgv.Cells(vFrm.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rowDgv.Cells(vFrm.ColPrimaryPO.Index).Value = .Primary_PO
                        rowDgv.Cells(vFrm.ColPrimaryGRN.Index).Value = .Primary_GRN
                        rowDgv.Cells(vFrm.ColPrimaryMasterUnits.Index).Value = .Primary_Master_Unit
                        rowDgv.Cells(vFrm.ColPrimaryUnit.Index).Value = .Primary_Units
                        rowDgv.Cells(vFrm.ColPrimarySecUnit.Index).Value = .Primary_Secondary_Units
                        rowDgv.Cells(vFrm.ColPrimaryCurrency.Index).Value = .Primary_Currency
                        rowDgv.Cells(vFrm.Col_AtAmount.Index).Value = ObjCD.DigitGrouping(.At_Amount, DefPrecBase)
                        rowDgv.Cells(vFrm.ColDiscFromPerct.Index).Value = .Disc_From_Perct
                        If .Primary_Units = .Primary_Units_Multi Then
                            rowDgv.Cells(vFrm.ColQtyMulti_Item.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, _
                            .Decimal_Product)
                        Else
                            rowDgv.Cells(vFrm.ColQtyMulti_Item.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, _
                            Me.GetDecimalMulti(TBLMultiUnit, .Primary_Product, .Primary_Units_Multi))
                        End If
                        rowDgv.Cells(vFrm.ColUnitsCodeMulti_Item.Index).Value = .Unit_Code_Multi
                        rowDgv.Cells(vFrm.ColRatioReference_Item.Index).Value = .Ratio_Reference
                        rowDgv.Cells(vFrm.ColRatioDefault_Item.Index).Value = .Ratio_Default
                        rowDgv.Cells(vFrm.ColPrimaryUnitsMulti_Item.Index).Value = .Primary_Units_Multi
                        rowDgv.Cells(vFrm.ColItem_Adjustment.Index).Value = ObjCD.DigitGrouping(.Adjustment, .Decimal_Currency)
                        rowDgv.Cells(vFrm.ColItem_AdjustmentReason.Index).Value = .Reason
                        rowDgv.Cells(vFrm.ColItem_AdjustUser.Index).Value = .User_Code
                        If .Adjustment_Timestamp = #1/1/1950# Then
                            rowDgv.Cells(vFrm.ColItemAdjustmentTimestamp.Index).Value = "-"
                        Else
                            rowDgv.Cells(vFrm.ColItemAdjustmentTimestamp.Index).Value = Format(.Adjustment_Timestamp, "dd MMM yyyy HH:mm")
                        End If

                        If .Primary_Units <> .Primary_Units_Multi Then
                            rowDgv.Cells(vFrm.ColRatioDesc_Item.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, _
                                    Me.GetDecimalMulti(TBLMultiUnit, .Primary_Product, .Primary_Units_Multi)) _
                                    & " " & .Unit_Code_Multi & " = " & ObjCD.DigitGrouping(.Ratio_Default, _
                                    .Decimal_Product) & " " & .Units
                        Else
                            rowDgv.Cells(vFrm.ColRatioDesc_Item.Index).Value = "-"
                        End If

                        rowDgv.Cells(vFrm.ColPriceByMultiUnit.Index).Value = .price_by_multi_unit
                        If .price_by_multi_unit = 1 Then
                            rowDgv.Cells(vFrm.ColPriceByMultiUnit_str.Index).Value = "/" + .Unit_Code_Multi
                        Else
                            rowDgv.Cells(vFrm.ColPriceByMultiUnit_str.Index).Value = "/" + .Units
                        End If
                        rowDgv.Cells(vFrm.ColATCost_Unit.Index).Value = "/" + rowDgv.Cells(vFrm.ColUnitsCodeMulti_Item.Index).Value
                    End With
                    'With row
                    '    vDGVItem.Rows.Add((vDGVItem.Rows.Count + 1).ToString, .Product_Code, .Description, ._Supplier_sProductCode_, ._Supplier_sProductDesc_, _
                    '    .PO_No, .GRN_No, .Stock_In_Purchase, .Supplier_Do_No, .Qty_Master, .Qty_Per_Master, .Master_Unit, .Qty, _
                    '    ObjCD.DigitGrouping(.Total_Qty, .Decimal_Product), .Units, .BatchNumber, .Secondary_Qty, .Secondary_Unit, .Currency_Code, ObjCD.DigitGrouping(.Unit_Price, .Decimal_Currency), _
                    '    ObjCD.DigitGrouping(.Discount, 2), ObjCD.DigitGrouping(.Disc_Amount, .Decimal_Currency), ObjCD.DigitGrouping(.Nett_Price, .Decimal_Currency), _
                    '    ObjCD.DigitGrouping(.Amount, .Decimal_Currency), ObjCD.DigitGrouping(.At_Price, DefPrecBase), "Show", .Remarks, _
                    '    .Primary_Product, .Primary_PO, .Primary_GRN, .Primary_Master_Unit, .Primary_Units, _
                    '    .Primary_Secondary_Units, .Primary_Currency, ObjCD.DigitGrouping(.At_Amount, DefPrecBase), .Disc_From_Perct)
                    'End With
                Next

                row = TBL.Rows(0)
                vPrimaryCurrencyRujukan = row.Primary_Currency
                CurrencyCOdeRujukan = row.Currency_Code
            End If
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'isikan tabel item dp
            If DTA1.FillBy_PrimaryAOP(TBL1, vPrimaryTransaksi) > 0 Then
                For Each row1 In TBL1.Rows
                    With row1
                        If .Tipe_DP = 1 Then 'aul add
                            If Me.GetManualDP_Data(.Primary_AOP_DP, row1) = False Then
                                'jika false artinya balance manual tidak ditemukan
                                GoTo SkipInsertToDgv
                            End If
                        End If
                        vDGVDP.Rows.Add("0", .No_AOP_DP, .Currency_Code, ObjCD.DigitGrouping(.Sub_Total, .Default_Precision), .Remarks, _
                        .Primary_AOP_DP, .Tipe_DP)
SkipInsertToDgv:
                    End With
                Next
            End If

            'isikan tabel item exchange rate
            If vCalculatedBy = ClsTransaksiAOP.EnumCalculatesBy.AOP Then
                Dim DTAEX As New DseTransaksiAOPTableAdapters.DataTable3TableAdapter
                Dim TBLEX As New DseTransaksiAOP.DataTable3DataTable
                Dim RowEX As DseTransaksiAOP.DataTable3Row
                If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
                    For Each RowEX In TBLEX.Rows
                        With RowEX
                            vDGVExchangeRate.Rows.Add("0", .NoTransaksiAOP, CurrencyCOdeRujukan, _
                            .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, vPrimaryCurrencyRujukan)
                        End With
                    Next
                End If
                DTAEX.Dispose()
                TBLEX.Dispose()
            ElseIf vCalculatedBy = ClsTransaksiAOP.EnumCalculatesBy.GRN Then
                Dim DTAEX As New DseTransaksiAOPTableAdapters.DataTable4TableAdapter
                Dim TBLEX As New DseTransaksiAOP.DataTable4DataTable
                Dim RowEX As DseTransaksiAOP.DataTable4Row
                If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
                    For Each RowEX In TBLEX.Rows
                        With RowEX
                            vDGVExchangeRate.Rows.Add("0", .GRN_No, CurrencyCOdeRujukan, _
                            .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, vPrimaryCurrencyRujukan)
                        End With
                    Next
                End If
                DTAEX.Dispose()
                TBLEX.Dispose()
            ElseIf vCalculatedBy = ClsTransaksiAOP.EnumCalculatesBy.PO Then
                Dim DTAEX As New DseTransaksiAOPTableAdapters.DataTable5TableAdapter
                Dim TBLEX As New DseTransaksiAOP.DataTable5DataTable
                Dim RowEX As DseTransaksiAOP.DataTable5Row
                If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
                    For Each RowEX In TBLEX.Rows
                        With RowEX
                            vDGVExchangeRate.Rows.Add("0", .PO_No, CurrencyCOdeRujukan, _
                            .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, vPrimaryCurrencyRujukan)
                        End With
                    Next
                End If
                DTAEX.Dispose()
                TBLEX.Dispose()
            End If

            'isikan dgv PR
            DTA2.FillBy_PrimaryAOP(TBL2, vPrimaryTransaksi)
            For Each row2 In TBL2.Rows
                vDGVPR.Rows.Add("", row2.PruchaseReturnNo, row2.GRN_No, row2.Remarks, row2.Primary_PR)
            Next

            'set nomer
            ObjCD.SetNumber(vDGVDP)
            ObjCD.SetNumber(vDGVExchangeRate)
            ObjCD.SetNumber(vDGVPR)
            ObjCD.SetNumber(vDGVCostItem)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
            DTA1.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
        End Try
    End Function

    Public Function GetManualDP_Data(ByVal vPrimaryAOP_DP As Double, ByRef row1 As DseTransaksiAOP.DataTable2Row) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tblbalance_purchase_dp1TableAdapter
        Dim Tbl As New DseTransaksiAOP.tblbalance_purchase_dp1DataTable
        Dim row As DseTransaksiAOP.tblbalance_purchase_dp1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(Tbl, AccModule.ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP, vPrimaryAOP_DP) > 0 Then
                row = Tbl.Rows(0)
                row1.No_AOP_DP = row.No_AOP_DP
                row1.Currency_Code = row.Currency_Code
                row1.Default_Precision = row.Default_Precision
                row1.Sub_Total = row.Sub_Total

                Return True
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
