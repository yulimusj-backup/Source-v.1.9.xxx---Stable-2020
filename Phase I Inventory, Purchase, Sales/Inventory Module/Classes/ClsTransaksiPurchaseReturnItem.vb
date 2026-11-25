Option Explicit On
Option Strict Off

Public Class ClsTransaksiPurchaseReturnItem

    '' Dipakai GRN
    Public Sub IsiDGVByPrimarySR(ByVal listPR As List(Of Double), ByVal vDGVItem As DataGridView)
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_itemview1TableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview1DataTable
        Dim row As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview1Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim vPrimaryTransaksi As Double = 0

        Try
            'kosongkan dgv
            vDGVItem.Rows.Clear()

            For Each vPrimaryTransaksi In listPR
                'isikan tabel
                If DTA.FillBy_PrimaryPR(TBL, vPrimaryTransaksi) > 0 Then
                    For Each row In TBL.Rows
                        With row
                            vDGVItem.Rows.Add("0", row.Product_Code, row.Description, _
                            row.GRN_No, row.Batch_No, row.Expiry_Date, row.Qty_Master, .Qty_Master, .MasterUnit, row.Qty, _
                            .Total_Qty, row.Unit_Code, row.Secondary_Qty, row.SecUnit, row.Remarks, row.Primary_Product, _
                            row.Primary_PO, row.Primary_Master_Units, row.Primary_Units, row.Primary_Secondary_Units)
                        End With
                    Next
                End If
            Next


            'set nomer
            ObjCD.SetNumber(vDGVItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
        End Try
    End Sub

    Public Function IsiDGV_Default(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_itemview2TableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview2DataTable
        Dim row As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview2Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjPO As New ClsTransaksiPO
        Dim PrimaryPO As Double = -1
        Dim PO_NO As String = ""

        Try
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    If row.Primary_PO <> PrimaryPO Then
                        PO_NO = ObjPO.GetNoTransaksiPO(row.Primary_PO)
                    End If
                    With row
                        Dim expiryDate As String
                        If row.Expiry_Date <> #1/1/1950# Then
                            expiryDate = row.Expiry_Date
                        Else
                            expiryDate = "-"
                        End If

                        vDGVItem.Rows.Add("0", row.Product_Code, row.Description, _
                       PO_NO, row.Batch_No, expiryDate, row.Qty_Master, .Qty_Master, .MasterUnit, row.Qty, _
                       ObjCD.DigitGrouping(.Total_Qty, ._Decimal), row.Unit_Code, row.Secondary_Qty, row.SecUnit, row.Remarks, row.Primary_Product, _
                       row.Primary_PO, row.Primary_Master_Units, row.Primary_Units, row.Primary_Secondary_Units)
                    End With
                Next
            End If
            'set nomer
            ObjCD.SetNumber(vDGVItem)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjPO = Nothing
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function IsiDGV_Multi(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_itemview2_MultiTableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview2_MultiDataTable
        Dim row As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview2_MultiRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjPO As New ClsTransaksiPO
        Dim PrimaryPO As Double = -1
        Dim PO_NO As String = ""
        Dim RowDgv As DataGridViewRow
        Dim i As New FrmTransaksiStockOutPurchaseReturn2
        Dim objProduct As New Core_Module.clsMasterProduct
        Try
            vDGVItem.Rows.Clear()
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    If row.Primary_PO <> PrimaryPO Then
                        PO_NO = ObjPO.GetNoTransaksiPO(row.Primary_PO)
                    End If
                    Dim expiryDate As String
                    If row.Expiry_Date <> #1/1/1950# Then
                        expiryDate = row.Expiry_Date
                    Else
                        expiryDate = "-"
                    End If

                    ''@rosadinnnn
                    ''@MUaT
                    vDGVItem.Rows.Add()
                    RowDgv = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                    With RowDgv
                        .Cells(i.ColPONODGVPR.Index).Value = ""
                        .Cells(i.colCodeNoDGVPR.Index).Value = row.Product_Code
                        .Cells(i.colDescDGVPR.Index).Value = row.Description
                        .Cells(i.ColPONODGVPR.Index).Value = PO_NO
                        .Cells(i.ColBatchNoDGVPR.Index).Value = row.Batch_No
                        .Cells(i.ColExpiryDateDGVPR.Index).Value = expiryDate
                        .Cells(i.ColQtyMasterDGVPR.Index).Value = row.Qty_Master
                        .Cells(i.ColQtyPerMasterDGVPR.Index).Value = row.Qty_Per_Master
                        .Cells(i.ColMasterUnitsDGVPR.Index).Value = row.MasterUnit
                        .Cells(i.colQTYDGVPR.Index).Value = row.Qty
                        .Cells(i.ColTotalQtyDGVPR.Index).Value = ObjCD.DigitGrouping(row.Total_Qty, row._Decimal)
                        .Cells(i.colUnitDGVPR.Index).Value = row.Unit_Code
                        .Cells(i.ColSecQtyDGVPR.Index).Value = row.Secondary_Qty
                        .Cells(i.ColSecUnitDGVPR.Index).Value = row.SecUnit
                        .Cells(i.colRemarksDGVPR.Index).Value = row.Remarks
                        .Cells(i.colPrimaryCodeNoDGVPR.Index).Value = row.Primary_Product
                        .Cells(i.ColPrimaryPODGVPR.Index).Value = row.Primary_PO
                        .Cells(i.ColPrimaryMasterUnitsDGVPR.Index).Value = row.Primary_Master_Units
                        .Cells(i.ColPrimaryUnitDGVPR.Index).Value = row.Primary_Units
                        .Cells(i.ColPrimarySecUnitDGVPR.Index).Value = row.Primary_Secondary_Units
                        .Cells(i.ColQtyMulti_PR.Index).Value = ObjCD.DigitGrouping(row.Total_Qty_Multi, objProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi))
                        .Cells(i.ColUnitCodeMulti_PR.Index).Value = row.Unit_Code_Multi
                        If row.Primary_Unit_Multi <> row.Primary_Units Then
                            .Cells(i.ColRatioDesc_PR.Index).Value = ObjCD.DigitGrouping(row.Ratio_Reference, _
                                    objProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi)) _
                                    & " " & row.Unit_Code_Multi & " = " & ObjCD.DigitGrouping(row.Ratio_Default, _
                                    objProduct.GetAngkaDecimal(row.Primary_Product)) & " " & row.Unit_Code
                        Else
                            .Cells(i.ColRatioDesc_PR.Index).Value = "-"
                        End If
                        .Cells(i.ColRaioReference_PR.Index).Value = row.Ratio_Reference
                        .Cells(i.ColRatioDefault_Pr.Index).Value = row.Ratio_Default
                        .Cells(i.ColPrimaryUnitsMulti_Pr.Index).Value = row.Primary_Unit_Multi
                    End With
                Next
            End If
            ObjCD.SetNumber(vDGVItem)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjPO = Nothing
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItem As DataGridView, _
  ByVal vPrimaryTransaksi As Double) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_itemTableAdapter
        Dim objProduct As New Core_Module.clsMasterProduct
        Dim objitem As New ClsTransaksiPOItem
        Try
            For Each row In vdgvItem.Rows
                'simpan ke basis data item
                Dim Expirydate As Date
                If IsDate(row.Cells("ColExpiryDate").Value) = True Then
                    Expirydate = row.Cells("ColExpiryDate").Value
                Else
                    Expirydate = #1/1/1950#
                End If
                DTA.Insert(vPrimaryTransaksi, row.Cells("ColPrimaryCodeNo").Value, row.Cells("ColPrimaryPO").Value, row.Cells("ColBatchNo").Value, _
                Expirydate, row.Cells("ColQtyMaster").Value, row.Cells("ColQtyPerMaster").Value, _
                row.Cells("ColPrimaryMasterUnits").Value, row.Cells("ColQty").Value, _
                row.Cells("ColTotalQty").Value, row.Cells("ColPrimaryUnit").Value, row.Cells("ColSecQty").Value, _
                row.Cells("ColPrimarySecUnit").Value, row.Cells("colRemarks").Value)
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            objProduct = Nothing
            objitem = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_itemTableAdapter
        Try
            'hapus data item
            DTA.Delete(vPrimaryTransaksi)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
