Option Explicit On
Option Strict Off

Public Class ClsTransaksiQuotationItem
    Public Function IsiDGV(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
   ByVal vDT_ItemDelivery As DseTransaksiQuotation.tbltrans_quotation_itemdeliveryDataTable, ByVal vDefPrecCurQuo As Integer) As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotation_item_view1_MultiTableAdapter
        Dim TBL As New DseTransaksiQuotation.tbltrans_quotation_item_view1_MultiDataTable
        Dim row As DseTransaksiQuotation.tbltrans_quotation_item_view1_MultiRow
        Dim FrmQuo As New FrmTransaksiQuotation
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim DeliveryDate_Str As String = ""

        Dim DefPrecMulti As Integer = 0
        Dim DefPrecDefault As Integer = 0
        Dim ObjProduct As New Core_Module.clsMasterProduct

        Dim DTAItemDelivery As New DseTransaksiQuotationTableAdapters.tbltrans_quotation_itemdeliveryTableAdapter
        'Dim rowItemDelivery As DseTransaksiQuotation.tbltrans_quotation_itemdeliveryRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemDelivery.Connection.ConnectionString = DTA.Connection.ConnectionString

            'item(delivery)
            DTAItemDelivery.FillBy_PrimaryTrans(vDT_ItemDelivery, vPrimaryTransaksi)

            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        'vDGVItem.Rows.Add("0", .Product_Code, .Description, ._Customer_s_Code_No, ._Customer_s_Desc, _
                        'ObjCD.DigitGrouping(.Qty, ._Decimal), .Unit_Code, .Unit_Price, .Discount, .Nett_Price, .Amount, CBool(.Delivery_Date_Cont), _
                        '.Delivery_Date, .Remarks, .Primary_Product, .Primary_Units, CBool(.Use_Secondary))
                        If CBool(.Delivery_Date_Cont) = True Then
                            If .Delivery_Date = #1/1/1950# Then
                                DeliveryDate_Str = "Multiple"
                            Else
                                DeliveryDate_Str = Format(.Delivery_Date, "dd MMM yyyy")
                            End If
                        Else
                            DeliveryDate_Str = "-"
                        End If
                        'FrmQuo.dgvItem.Rows.Add()
                        DefPrecMulti = ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)
                        DefPrecDefault = ObjProduct.GetAngkaDecimal(.Primary_Product)
                        vDGVItem.Rows.Add()
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(FrmQuo.ColNo.Index).Value = vDGVItem.Rows.Count
                        rowDGV.Cells(FrmQuo.ColCodeNo.Index).Value = .Product_Code
                        rowDGV.Cells(FrmQuo.ColDescription.Index).Value = .Description
                        rowDGV.Cells(FrmQuo.ColSupplierCodeNo.Index).Value = ._Customer_s_Code_No
                        rowDGV.Cells(FrmQuo.ColSuppDesc.Index).Value = ._Customer_s_Desc
                        rowDGV.Cells(FrmQuo.ColQty.Index).Value = ObjCD.DigitGrouping(.Qty, ._Decimal)
                        rowDGV.Cells(FrmQuo.ColUnit.Index).Value = .Unit_Code
                        rowDGV.Cells(FrmQuo.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(.Unit_Price, vDefPrecCurQuo)
                        rowDGV.Cells(FrmQuo.ColDiscount.Index).Value = ObjCD.DigitGrouping(.Discount, 2)
                        rowDGV.Cells(FrmQuo.ColNettPrice.Index).Value = ObjCD.DigitGrouping(.Nett_Price, vDefPrecCurQuo)
                        rowDGV.Cells(FrmQuo.ColAmount.Index).Value = ObjCD.DigitGrouping(.Amount, vDefPrecCurQuo)
                        rowDGV.Cells(FrmQuo.ColDeliveryCont.Index).Value = CBool(.Delivery_Date_Cont)
                        rowDGV.Cells(FrmQuo.ColDeliveryDate.Index).Value = DeliveryDate_Str
                        rowDGV.Cells(FrmQuo.ColRemark.Index).Value = .Remarks
                        rowDGV.Cells(FrmQuo.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rowDGV.Cells(FrmQuo.ColPrimaryUnits.Index).Value = .Primary_Units
                        rowDGV.Cells(FrmQuo.ColUseSecondary.Index).Value = CBool(.Use_Secondary)
                        rowDGV.Cells(FrmQuo.ColDiscAmount.Index).Value = ObjCD.DigitGrouping(.Disc_Amount, vDefPrecCurQuo)
                        rowDGV.Cells(FrmQuo.ColDiscFromPerct.Index).Value = .Disc_From_Perct
                        rowDGV.Cells(FrmQuo.ColIdentifier.Index).Value = .Identifier

                        rowDGV.Cells(FrmQuo.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(.Qty_Multi, DefPrecMulti)
                        rowDGV.Cells(FrmQuo.ColUnitsMulti.Index).Value = .Units_Code_Multi
                        rowDGV.Cells(FrmQuo.ColPrimaryUnitsMulti.Index).Value = .Primary_Units_Multi
                        rowDGV.Cells(FrmQuo.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(.Ratio_Default, DefPrecDefault)
                        rowDGV.Cells(FrmQuo.ColRatioReference.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti)
                        If .Primary_Units <> .Primary_Units_Multi Then
                            rowDGV.Cells(FrmQuo.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti) & " " & .Units_Code_Multi & "=" & ObjCD.DigitGrouping(.Ratio_Default, DefPrecDefault) & " " & .Unit_Code
                        Else
                            rowDGV.Cells(FrmQuo.ColRatioDesc.Index).Value = "-"
                        End If
                        rowDGV.Cells(FrmQuo.ColToleranceDown.Index).Value = ObjCD.DigitGrouping(.Tolerance_Down, DefPrecMulti)
                        rowDGV.Cells(FrmQuo.ColToleranceDownPerc.Index).Value = ObjCD.DigitGrouping(.Tolerance_Down_Perc, 2)
                        rowDGV.Cells(FrmQuo.ColToleranceDownFromPerc.Index).Value = ObjCD.DigitGrouping(.Tolerance_Down_From_Perc, 0)
                        rowDGV.Cells(FrmQuo.ColToleranceUp.Index).Value = ObjCD.DigitGrouping(.Tolerance_Up, DefPrecMulti)
                        rowDGV.Cells(FrmQuo.ColToleranceUpPerc.Index).Value = ObjCD.DigitGrouping(.Tolerance_Up_Perc, 2)
                        rowDGV.Cells(FrmQuo.ColToleranceUpFromPerc.Index).Value = ObjCD.DigitGrouping(.Tolerance_Up_From_Perc, 0)
                        rowDGV.Cells(FrmQuo.ColPriceByMultiUnit.Index).Value = .price_by_multi_unit
                        If .price_by_multi_unit = 1 Then
                            rowDGV.Cells(FrmQuo.ColPriceByMultiUnit_Str.Index).Value = "/" + .Units_Code_Multi
                        Else
                            rowDGV.Cells(FrmQuo.ColPriceByMultiUnit_Str.Index).Value = "/" + .Unit_Code
                        End If
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
            DTAItemDelivery.Dispose()
            TBL = Nothing
            ObjCD = Nothing
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItem As DataGridView, _
    ByVal vPrimaryTransaksi As Double, ByVal i As FrmTransaksiQuotation, ByVal vDT_ItemDelivery As DseTransaksiQuotation.tbltrans_quotation_itemdeliveryDataTable) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotation_itemTableAdapter
        Dim DelDateCont, UseSecondary As Byte
        Dim DelDate As Date
        Dim DTA_ItemDelivery As New DseTransaksiQuotationTableAdapters.tbltrans_quotation_itemdeliveryTableAdapter
        Dim rowItemDelivery As DseTransaksiQuotation.tbltrans_quotation_itemdeliveryRow

        Try
            For Each row In vdgvItem.Rows
                'simpan ke basis data item
                If row.Cells(i.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                    DelDateCont = 1
                    If IsDate(row.Cells(i.ColDeliveryDate.Index).Value) = True Then
                        DelDate = row.Cells(i.ColDeliveryDate.Index).Value
                    Else
                        DelDate = #1/1/1950#
                    End If
                Else
                    DelDateCont = 0
                    DelDate = #1/1/1950#
                End If
                If row.Cells(i.ColUseSecondary.Index).Value = Boolean.TrueString Then
                    UseSecondary = 1
                Else
                    UseSecondary = 0
                End If

                If row.Cells(i.ColIdentifier.Index).Value.ToString.StartsWith("-1") Then
                    row.Cells(i.ColIdentifier.Index).Value = vPrimaryTransaksi.ToString + row.Cells(i.ColIdentifier.Index).Value.ToString.Substring(2)
                End If
                DTA.Insert(vPrimaryTransaksi, row.Cells(i.ColPrimaryCodeNo.Index).Value, row.Cells(i.ColQty.Index).Value, row.Cells(i.ColPrimaryUnits.Index).Value, _
                DelDateCont, DelDate, row.Cells(i.ColUnitPrice.Index).Value, row.Cells(i.ColDiscount.Index).Value, row.Cells(i.ColNettPrice.Index).Value, _
                row.Cells(i.ColAmount.Index).Value, UseSecondary, row.Cells(i.ColRemark.Index).Value, row.Cells(i.ColIdentifier.Index).Value, _
                row.Cells(i.ColDiscAmount.Index).Value, row.Cells(i.ColDiscFromPerct.Index).Value, _
                row.Cells(i.ColQtyMulti.Index).Value, row.Cells(i.ColPrimaryUnitsMulti.Index).Value, row.Cells(i.ColRatioDefault.Index).Value, _
                row.Cells(i.ColRatioReference.Index).Value, row.Cells(i.ColToleranceDown.Index).Value, row.Cells(i.ColToleranceDownPerc.Index).Value, _
                row.Cells(i.ColToleranceDownFromPerc.Index).Value, row.Cells(i.ColToleranceUp.Index).Value, row.Cells(i.ColToleranceUpPerc.Index).Value, _
                row.Cells(i.ColToleranceUpFromPerc.Index).Value, row.Cells(i.ColPriceByMultiUnit.Index).Value)
            Next

            For Each rowItemDelivery In vDT_ItemDelivery.Rows
                With rowItemDelivery
                    If .identifier.StartsWith("-1") Then
                        .identifier = vPrimaryTransaksi.ToString + .identifier.Substring(2)
                    End If
                    DTA_ItemDelivery.InsertQuery(vPrimaryTransaksi, .identifier, .delivery_qty, .delivery_date, .remarks, .delivery_priority, .delivery_qty_multi)
                End With
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotation_itemTableAdapter
        Dim DTA_ItemDelivery As New DseTransaksiQuotationTableAdapters.tbltrans_quotation_itemdeliveryTableAdapter
        Try
            'hapus data item
            DTA.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_ItemDelivery.Connection.ConnectionString = DTA.Connection.ConnectionString

            'hapus data item
            DTA.Delete(vPrimaryTransaksi)
            DTA_ItemDelivery.DeleteQuery(vPrimaryTransaksi)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
