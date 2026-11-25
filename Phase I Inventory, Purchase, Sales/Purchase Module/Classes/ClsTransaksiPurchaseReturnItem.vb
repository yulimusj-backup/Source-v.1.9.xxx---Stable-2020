Option Explicit On
Option Strict Off

Public Class ClsTransaksiPurchaseReturnItem
    Public Sub IsikanTBLBatchNo(ByVal vTBL As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item5DataTable, ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double)
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item5TableAdapter

        Try
            DTA.Fill(vTBL, vPrimaryProduct, vPrimaryTransaksi)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLLocation(ByVal vTBL As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item6DataTable, ByVal vPrimaryProduct As Double, _
    ByVal vPrimaryPR As Double, ByVal vBatchNumber As String)
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item6TableAdapter

        Try
            DTA.Fill(vTBL, vPrimaryPR, vPrimaryProduct, vBatchNumber)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetJml3(ByVal vPrimaryPR As Double, ByVal vPrimaryCodeNo As Double, ByVal vPrimaryPO As Double) As Decimal
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item2TableAdapter

        Try
            If DTA.ScalarQuery_GetJml4(vPrimaryCodeNo, vPrimaryPO, vPrimaryPR).HasValue = True Then
                Return DTA.ScalarQuery_GetJml4(vPrimaryCodeNo, vPrimaryPO, vPrimaryPR).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml2(ByVal vPrimaryPR As Double, ByVal vPrimarCodeNo As Double, ByVal vBatchNo As String, ByVal vPrimaryLocation As Double) As Decimal
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item2TableAdapter

        Try
            If DTA.ScalarQuery_GetJml3(vPrimaryPR, vPrimarCodeNo, vBatchNo, vPrimaryLocation).HasValue = True Then
                Return DTA.ScalarQuery_GetJml3(vPrimaryPR, vPrimarCodeNo, vBatchNo, vPrimaryLocation).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vPrimaryPR As Double, ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double, _
    ByVal vPrimaryGRN As Double) As Decimal
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item2TableAdapter

        Try
            If DTA.ScalarQuery_GetJml2(vPrimaryGRN, vPrimaryProduct, vPrimaryPO, vPrimaryPR).HasValue = True Then
                Return DTA.ScalarQuery_GetJml2(vPrimaryGRN, vPrimaryProduct, vPrimaryPO, vPrimaryPR).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetReturnBalance(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double, _
    ByVal vPrimaryStockIn As Double, ByVal vPrimaryGRN As Double) As Decimal
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.PR_OBTableAdapter
        Dim DTA1 As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item2TableAdapter
        Dim QtyStockIn, QtyReturn As Decimal

        Try
            If DTA.ScalarQuery_GetJml(vPrimaryStockIn, vPrimaryProduct, vPrimaryPO).HasValue = True Then
                QtyStockIn = DTA.ScalarQuery_GetJml(vPrimaryStockIn, vPrimaryProduct, vPrimaryPO).Value
            Else
                QtyStockIn = 0
            End If
            If DTA1.ScalarQuery_GetJml(vPrimaryGRN, vPrimaryPO, vPrimaryProduct).HasValue = True Then
                QtyReturn = DTA1.ScalarQuery_GetJml(vPrimaryGRN, vPrimaryPO, vPrimaryProduct).Value
            Else
                QtyReturn = 0
            End If

            Return QtyStockIn - QtyReturn
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
        End Try
    End Function

    '' Dipakai GRN dan AOP ketika melakukan penambahan item return
    Public Sub IsiDGVByPrimaryPR_Default(ByVal listPR As List(Of Double), ByVal vDGVItem As DataGridView)
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
                If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                    For Each row In TBL.Rows
                        With row
                            vDGVItem.Rows.Add("0", row.Product_Code, row.Description, _
                            row.PO_No, row.GRN_No, row.PR_Number, row.Batch_No, row.Expiry_Date, row.Qty_Master, .Qty_Master, .MasterUnit, row.Qty, _
                            ObjCD.DigitGrouping(.Total_Qty, ._Decimal), row.Unit_Code, row.Secondary_Qty, row.SecUnit, row.Remarks, row.Primary_Product, _
                            row.Primary_PO, row.Primary_Master_Units, row.Primary_Units, row.Primary_Secondary_Units, row.Primary_GRN)
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

    ''@rosadinnn
    ''@MUaT
    Public Sub IsiDGVByPrimaryPR_Multi(ByVal listPR As List(Of Double), ByVal vDGVItem As DataGridView)
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_itemview1_MultiTableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview1_MultiDataTable
        Dim row As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview1_MultiRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim vPrimaryTransaksi As Double = 0
        Dim rowDgv As DataGridViewRow
        Dim i As New FrmTransaksiAOP2
        Dim objProduct As New Core_Module.clsMasterProduct

        Try
            'kosongkan dgv
            vDGVItem.Rows.Clear()

            For Each vPrimaryTransaksi In listPR
                'isikan tabel
                If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                    For Each row In TBL.Rows
                        'vDGVItem.Rows.Add("0", row.Product_Code, row.Description, _
                        'row.PO_No, row.GRN_No, row.PR_Number, row.Batch_No, row.Expiry_Date, row.Qty_Master, .Qty_Master, .MasterUnit, row.Qty, _
                        'ObjCD.DigitGrouping(.Total_Qty, ._Decimal), row.Unit_Code, row.Secondary_Qty, row.SecUnit, row.Remarks, row.Primary_Product, _
                        'row.Primary_PO, row.Primary_Master_Units, row.Primary_Units, row.Primary_Secondary_Units, row.Primary_GRN)
                        vDGVItem.Rows.Add()
                        rowDgv = vDGVItem.Rows(vDGVItem.RowCount - 1)
                        With rowDgv
                            .Cells(i.ColNo.Index).Value = ""
                            .Cells(i.ColCodeNo_PR.Index).Value = row.Product_Code
                            .Cells(i.ColDesc_PR.Index).Value = row.Description
                            .Cells(i.ColPONo_PR.Index).Value = row.PO_No
                            .Cells(i.ColGRNNo_PR.Index).Value = row.Stock_In_No
                            .Cells(i.ColPRNumber.Index).Value = row.PR_Number
                            .Cells(i.ColBatchNo_PR.Index).Value = row.Batch_No
                            .Cells(i.ColExpiryDate_PR.Index).Value = Format(row.Expiry_Date, "dd MMM yyyy")
                            '.Cells(i.QtyMaster_PR.Index).Value = row.Qty_Master
                            '.Cells(i.ColMasterUnit_PR.Index).Value = row.MasterUnit
                            .Cells(i.ColQty_PR.Index).Value = row.Qty
                            .Cells(i.ColTotalQty_PR.Index).Value = ObjCD.DigitGrouping(row.Total_Qty, row._Decimal)
                            .Cells(i.ColUnit_PR.Index).Value = row.Unit_Code
                            .Cells(i.ColSecQty_PR.Index).Value = row.Secondary_Qty
                            .Cells(i.ColSecUnit_PR.Index).Value = row.SecUnit
                            .Cells(i.ColRemarks_PR.Index).Value = row.Remarks
                            .Cells(i.ColPrimaryCodeNo_PR.Index).Value = row.Primary_Product
                            .Cells(i.ColPrimaryPR_Item.Index).Value = row.Primary_PO
                            .Cells(i.ColPrimaryMasterUnit_PR.Index).Value = row.Primary_Master_Units
                            .Cells(i.ColPrimaryUnits_PR.Index).Value = row.Primary_Units
                            .Cells(i.ColPrimarySecUnits_PR.Index).Value = row.Primary_Secondary_Units
                            .Cells(i.ColPrimaryGRN_PR.Index).Value = row.Primary_GRN
                            .Cells(i.ColQtyMulti_PR.Index).Value = ObjCD.DigitGrouping(row.Total_Qty_Multi, objProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi))
                            .Cells(i.ColUnitsCodeMulti_PR.Index).Value = row.Unit_Code_Multi
                            .Cells(i.ColRatioReference_PR.Index).Value = row.Ratio_Reference
                            .Cells(i.ColRatioDefault_PR.Index).Value = row.Ratio_Default
                            If row.Primary_Units <> row.Primary_Unit_Multi Then
                                .Cells(i.ColRatioDesc_PR.Index).Value = ObjCD.DigitGrouping(row.Ratio_Reference, _
                                        objProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi)) _
                                        & " " & row.Unit_Code_Multi & " = " & ObjCD.DigitGrouping(row.Ratio_Default, _
                                        objProduct.GetAngkaDecimal(row.Primary_Product)) & " " & row.Unit_Code
                            Else
                                .Cells(i.ColRatioDesc_PR.Index).Value = "-"
                            End If
                            .Cells(i.ColPrimaryUnitsMulti_Pr.Index).Value = row.Primary_Unit_Multi

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

    'Public Function IsiDGV_Default(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    'ByVal vFrm As FrmTransaksiPurchaseReturn) As Boolean
    '    Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_itemview2TableAdapter
    '    Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview2DataTable
    '    Dim row As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview2Row
    '    Dim ObjCD As New Core_Module.ClsCekData
    '    Dim rowDGV As DataGridViewRow
    '    Dim RoundingItem As Integer = 0
    '    Dim PrimaryPO As Double = -1
    '    Dim PO_Number As String = ""
    '    Dim ObjPO As New ClsTransaksiPO

    '    Try
    '        vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '        vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

    '        'kosongkan dgv
    '        vDGVItem.Rows.Clear()

    '        PrimaryPO = -1
    '        PO_Number = ""

    '        'isikan tabel
    '        If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
    '            row = TBL.Rows(0)
    '            RoundingItem = vFrm.ObjCurrency.GetDefaultPrecission(row.Primary_Currency)
    '            For Each row In TBL.Rows
    '                With row
    '                    If .Primary_PO <> PrimaryPO Then
    '                        PrimaryPO = .Primary_PO
    '                        PO_Number = ObjPO.GetNoTransaksiPO(.Primary_PO)
    '                    End If
    '                    .PO_No = PO_Number

    '                    Dim expiryDate As String
    '                    If row.Expiry_Date <> #1/1/1950# Then
    '                        expiryDate = row.Expiry_Date
    '                    Else
    '                        expiryDate = "-"
    '                    End If
    '                    vDGVItem.Rows.Add("0", row.Product_Code, row.Description, _
    '                   row.PO_No, row.Batch_No, expiryDate, row.Qty_Master, .Qty_Master, .MasterUnit, row.Qty, _
    '                   ObjCD.DigitGrouping(.Total_Qty, ._Decimal), row.Unit_Code, row.Secondary_Qty, row.SecUnit, row.Remarks, row.Primary_Product, _
    '                   row.Primary_PO, row.Primary_Master_Units, row.Primary_Units, row.Primary_Secondary_Units)

    '                    rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
    '                    rowDGV.Cells(vFrm.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Unit_Price, RoundingItem))
    '                    rowDGV.Cells(vFrm.ColDiscount.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Discount, 2))
    '                    rowDGV.Cells(vFrm.ColNettPrice.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Nett_Price, RoundingItem))
    '                    rowDGV.Cells(vFrm.ColAmount.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Amount, RoundingItem))
    '                    rowDGV.Cells(vFrm.ColPrimaryCurrency.Index).Value = .Primary_Currency
    '                    rowDGV.Cells(vFrm.ColCurrency.Index).Value = vFrm.ObjCurrency.GetCurrencyCode(.Primary_Currency)
    '                    rowDGV.Cells(vFrm.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(ObjCD.RemoveDecimalZero(CDec(.Exchanged_Rate)))
    '                    rowDGV.Cells(vFrm.ColPurchaseInIDR.Index).Value = .Purchase_Price_In_IDR
    '                End With
    '            Next
    '        End If
    '        'set nomer
    '        ObjCD.SetNumber(vDGVItem)

    '        vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '        vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

    '        'nilai balikan
    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '        TBL = Nothing
    '        ObjCD = Nothing
    '        ObjPO = Nothing
    '    End Try
    'End Function

    Public Function IsiDGV_Multi(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal vFrm As FrmTransaksiPurchaseReturn) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_itemview2_MultiTableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview2_MultiDataTable
        Dim row As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview2_MultiRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim RoundingItem As Integer = 0
        Dim PrimaryPO As Double = -1
        Dim PO_Number As String = ""
        Dim ObjPO As New ClsTransaksiPO
        Dim objProduct As New Core_Module.clsMasterProduct
        Dim DefPrecMulti As Integer = 0

        Try
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'kosongkan dgv
            vDGVItem.Rows.Clear()

            PrimaryPO = -1
            PO_Number = ""

            'isikan tabel
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                row = TBL.Rows(0)
                RoundingItem = vFrm.ObjCurrency.GetDefaultPrecission(row.Primary_Currency)
                For Each row In TBL.Rows
                    With row
                        If .Primary_PO <> PrimaryPO Then
                            PrimaryPO = .Primary_PO
                            PO_Number = ObjPO.GetNoTransaksiPO(.Primary_PO)
                        End If
                        .PO_No = PO_Number

                        Dim expiryDate As String
                        If row.Expiry_Date <> #1/1/1950# Then
                            expiryDate = row.Expiry_Date
                        Else
                            expiryDate = "-"
                        End If
                        ' vDGVItem.Rows.Add("0", row.Product_Code, row.Description, _
                        'row.PO_No, row.Batch_No, expiryDate, row.Qty_Master, .Qty_Master, .MasterUnit, row.Qty, _
                        'ObjCD.DigitGrouping(.Total_Qty, ._Decimal), row.Unit_Code, row.Secondary_Qty, row.SecUnit, row.Remarks, row.Primary_Product, _
                        'row.Primary_PO, row.Primary_Master_Units, row.Primary_Units, row.Primary_Secondary_Units)

                        vDGVItem.Rows.Add()
                        DefPrecMulti = objProduct.GetAngkaDecimal(.Primary_Product, .Primary_Unit_Multi)

                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(vFrm.DataGridViewTextBoxColumn4.Index).Value = ""
                        rowDGV.Cells(vFrm.ColCodeNo.Index).Value = .Product_Code
                        rowDGV.Cells(vFrm.ColDescription.Index).Value = .Description
                        rowDGV.Cells(vFrm.ColPONO.Index).Value = .PO_No
                        rowDGV.Cells(vFrm.ColBatchNo.Index).Value = .Batch_No
                        rowDGV.Cells(vFrm.ColExpiryDate.Index).Value = expiryDate
                        rowDGV.Cells(vFrm.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, DefPrecMulti)
                        rowDGV.Cells(vFrm.ColUnitCodeMulti.Index).Value = .Unit_Code_Multi
                        If .Primary_Units <> .Primary_Unit_Multi Then
                            rowDGV.Cells(vFrm.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, _
                                    DefPrecMulti) _
                                    & " " & .Unit_Code_Multi & " = " & ObjCD.DigitGrouping(.Ratio_Default, _
                                    objProduct.GetAngkaDecimal(.Primary_Product)) & " " & .Unit_Code
                        Else
                            rowDGV.Cells(vFrm.ColRatioDesc.Index).Value = "-"
                        End If
                        rowDGV.Cells(vFrm.ColTotalQty.Index).Value = ObjCD.DigitGrouping(.Total_Qty, objProduct.GetAngkaDecimal(.Primary_Product))
                        rowDGV.Cells(vFrm.ColUnit.Index).Value = .Unit_Code
                        rowDGV.Cells(vFrm.ColQtyMaster.Index).Value = .Qty_Master
                        rowDGV.Cells(vFrm.ColQtyPerMaster.Index).Value = .Qty_Per_Master
                        rowDGV.Cells(vFrm.ColMasterUnits.Index).Value = .MasterUnit
                        rowDGV.Cells(vFrm.ColQty.Index).Value = .Qty
                        rowDGV.Cells(vFrm.ColSecQty.Index).Value = .Secondary_Qty
                        rowDGV.Cells(vFrm.ColSecUnit.Index).Value = .SecUnit
                        rowDGV.Cells(vFrm.colRemarks.Index).Value = .Remarks
                        rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rowDGV.Cells(vFrm.ColPrimaryPO.Index).Value = .Primary_PO
                        rowDGV.Cells(vFrm.ColPrimaryMasterUnits.Index).Value = .Primary_Master_Units
                        rowDGV.Cells(vFrm.ColPrimaryUnit.Index).Value = .Primary_Units
                        rowDGV.Cells(vFrm.ColPrimarySecUnit.Index).Value = .Primary_Secondary_Units

                        rowDGV.Cells(vFrm.ColCurrency.Index).Value = vFrm.ObjCurrency.GetCurrencyCode(.Primary_Currency)
                        rowDGV.Cells(vFrm.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Unit_Price, RoundingItem))
                        rowDGV.Cells(vFrm.ColDiscount.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Discount, 2))
                        rowDGV.Cells(vFrm.ColNettPrice.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Nett_Price, RoundingItem))
                        rowDGV.Cells(vFrm.ColAmount.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Amount, RoundingItem))
                        rowDGV.Cells(vFrm.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(ObjCD.RemoveDecimalZero(CDec(.Exchanged_Rate)))
                        rowDGV.Cells(vFrm.ColPrimaryCurrency.Index).Value = .Primary_Currency
                        rowDGV.Cells(vFrm.ColPurchaseInIDR.Index).Value = .Purchase_Price_In_IDR

                        rowDGV.Cells(vFrm.ColRatioReference.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti)
                        rowDGV.Cells(vFrm.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(.Ratio_Default, objProduct.GetAngkaDecimal(.Primary_Product))
                        rowDGV.Cells(vFrm.ColPrimaryUnitMulti.Index).Value = .Primary_Unit_Multi

                        rowDGV.Cells(vFrm.ColPriceByMultiUnit.Index).Value = .price_by_multi_unit
                        If .price_by_multi_unit = 1 Then
                            rowDGV.Cells(vFrm.ColPriceByMultiUnit_str.Index).Value = "/" + .Unit_Code_Multi
                        Else
                            rowDGV.Cells(vFrm.ColPriceByMultiUnit_str.Index).Value = "/" + .Unit_Code
                        End If
                    End With
                Next
            End If
            'set nomer
            ObjCD.SetNumber(vDGVItem)

            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

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

    '  Public Function IsiDGV_Default(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    'ByVal vFrm As FrmTransaksiPurchaseReturn_Quick) As Boolean
    '      Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_itemview2TableAdapter
    '      Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview2DataTable
    '      Dim row As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview2Row
    '      Dim ObjCD As New Core_Module.ClsCekData
    '      Dim rowDGV As DataGridViewRow
    '      Dim RoundingItem As Integer = 0
    '      Dim PrimaryPO As Double = -1
    '      Dim PO_Number As String = ""
    '      Dim ObjPO As New ClsTransaksiPO

    '      Try
    '          vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '          vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

    '          'kosongkan dgv
    '          vDGVItem.Rows.Clear()

    '          PrimaryPO = -1
    '          PO_Number = ""

    '          'isikan tabel
    '          If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
    '              row = TBL.Rows(0)
    '              RoundingItem = vFrm.ObjCurrency.GetDefaultPrecission(row.Primary_Currency)
    '              For Each row In TBL.Rows
    '                  With row
    '                      If .Primary_PO <> PrimaryPO Then
    '                          PrimaryPO = .Primary_PO
    '                          PO_Number = ObjPO.GetNoTransaksiPO(.Primary_PO)
    '                      End If
    '                      .PO_No = PO_Number

    '                      Dim expiryDate As String
    '                      If row.Expiry_Date <> #1/1/1950# Then
    '                          expiryDate = row.Expiry_Date
    '                      Else
    '                          expiryDate = "-"
    '                      End If
    '                      vDGVItem.Rows.Add()

    '                      rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
    '                      rowDGV.Cells(vFrm.ColCodeNo.Index).Value = .Product_Code
    '                      rowDGV.Cells(vFrm.ColDescription.Index).Value = .Description
    '                      rowDGV.Cells(vFrm.ColPONO.Index).Value = .PO_No
    '                      rowDGV.Cells(vFrm.ColBatchNo.Index).Value = .Batch_No
    '                      rowDGV.Cells(vFrm.ColExpiryDate.Index).Value = expiryDate
    '                      rowDGV.Cells(vFrm.ColQtyMaster.Index).Value = 0
    '                      rowDGV.Cells(vFrm.ColQtyPerMaster.Index).Value = 0
    '                      rowDGV.Cells(vFrm.ColMasterUnits.Index).Value = ""
    '                      rowDGV.Cells(vFrm.ColQty.Index).Value = .Qty
    '                      rowDGV.Cells(vFrm.ColTotalQty.Index).Value = ObjCD.DigitGrouping(.Total_Qty, ._Decimal)
    '                      rowDGV.Cells(vFrm.ColUnit.Index).Value = .Unit_Code
    '                      rowDGV.Cells(vFrm.ColSecQty.Index).Value = -1
    '                      rowDGV.Cells(vFrm.ColSecUnit.Index).Value = ""
    '                      rowDGV.Cells(vFrm.colRemarks.Index).Value = .Remarks
    '                      rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value = .Primary_Product
    '                      rowDGV.Cells(vFrm.ColPrimaryPO.Index).Value = .Primary_PO
    '                      rowDGV.Cells(vFrm.ColPrimaryMasterUnits.Index).Value = -1
    '                      rowDGV.Cells(vFrm.ColPrimaryUnit.Index).Value = .Primary_Units
    '                      rowDGV.Cells(vFrm.ColPrimarySecUnit.Index).Value = .Primary_Secondary_Units

    '                      rowDGV.Cells(vFrm.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Unit_Price, RoundingItem))
    '                      rowDGV.Cells(vFrm.ColDiscount.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Discount, 2))
    '                      rowDGV.Cells(vFrm.ColNettPrice.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Nett_Price, RoundingItem))
    '                      rowDGV.Cells(vFrm.ColAmount.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Amount, RoundingItem))
    '                      rowDGV.Cells(vFrm.ColPrimaryCurrency.Index).Value = .Primary_Currency
    '                      rowDGV.Cells(vFrm.ColCurrency.Index).Value = vFrm.ObjCurrency.GetCurrencyCode(.Primary_Currency)
    '                      rowDGV.Cells(vFrm.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(ObjCD.RemoveDecimalZero(CDec(.Exchanged_Rate)))
    '                      rowDGV.Cells(vFrm.ColPrimaryLocation.Index).Value = .primary_location
    '                      rowDGV.Cells(vFrm.ColLocation.Index).Value = .Storage_Area_Code
    '                      rowDGV.Cells(vFrm.ColNo.Index).Value = vDGVItem.Rows.Count
    '                      rowDGV.Cells(vFrm.ColPurchasInIDR.Index).Value = .Purchase_Price_In_IDR
    '                  End With
    '              Next
    '          End If

    '          vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

    '          'nilai balikan
    '          Return True
    '      Catch ex As Exception
    '          MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '      Finally
    '          DTA.Dispose()
    '          TBL = Nothing
    '          ObjCD = Nothing
    '          ObjPO = Nothing
    '      End Try
    '  End Function

    Public Function IsiDGV_Multi(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
            ByVal vFrm As FrmTransaksiPurchaseReturn_Quick) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_itemview2_MultiTableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview2_MultiDataTable
        Dim row As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_itemview2_MultiRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim RoundingItem As Integer = 0
        Dim PrimaryPO As Double = -1
        Dim PO_Number As String = ""
        Dim ObjPO As New ClsTransaksiPO
        Dim objProduct As New Core_Module.clsMasterProduct

        Try
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'kosongkan dgv
            vDGVItem.Rows.Clear()

            PrimaryPO = -1
            PO_Number = ""

            'isikan tabel
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                row = TBL.Rows(0)
                RoundingItem = vFrm.ObjCurrency.GetDefaultPrecission(row.Primary_Currency)
                For Each row In TBL.Rows
                    With row
                        If .Primary_PO <> PrimaryPO Then
                            PrimaryPO = .Primary_PO
                            PO_Number = ObjPO.GetNoTransaksiPO(.Primary_PO)
                        End If
                        .PO_No = PO_Number

                        Dim expiryDate As String
                        If row.Expiry_Date <> #1/1/1950# Then
                            expiryDate = row.Expiry_Date
                        Else
                            expiryDate = "-"
                        End If
                        vDGVItem.Rows.Add()

                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(vFrm.ColCodeNo.Index).Value = .Product_Code
                        rowDGV.Cells(vFrm.ColDescription.Index).Value = .Description
                        rowDGV.Cells(vFrm.ColPONO.Index).Value = .PO_No
                        rowDGV.Cells(vFrm.ColBatchNo.Index).Value = .Batch_No
                        rowDGV.Cells(vFrm.ColExpiryDate.Index).Value = expiryDate
                        rowDGV.Cells(vFrm.ColQtyMaster.Index).Value = 0
                        rowDGV.Cells(vFrm.ColQtyPerMaster.Index).Value = 0
                        rowDGV.Cells(vFrm.ColMasterUnits.Index).Value = ""
                        rowDGV.Cells(vFrm.ColQty.Index).Value = .Qty
                        rowDGV.Cells(vFrm.ColTotalQty.Index).Value = ObjCD.DigitGrouping(.Total_Qty, ._Decimal)
                        rowDGV.Cells(vFrm.ColUnit.Index).Value = .Unit_Code
                        rowDGV.Cells(vFrm.ColSecQty.Index).Value = -1
                        rowDGV.Cells(vFrm.ColSecUnit.Index).Value = ""
                        rowDGV.Cells(vFrm.colRemarks.Index).Value = .Remarks
                        rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rowDGV.Cells(vFrm.ColPrimaryPO.Index).Value = .Primary_PO
                        rowDGV.Cells(vFrm.ColPrimaryMasterUnits.Index).Value = -1
                        rowDGV.Cells(vFrm.ColPrimaryUnit.Index).Value = .Primary_Units
                        rowDGV.Cells(vFrm.ColPrimarySecUnit.Index).Value = .Primary_Secondary_Units

                        rowDGV.Cells(vFrm.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Unit_Price, RoundingItem))
                        rowDGV.Cells(vFrm.ColDiscount.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Discount, 2))
                        rowDGV.Cells(vFrm.ColDiscAmount.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Disc_Amount, RoundingItem))
                        rowDGV.Cells(vFrm.ColDiscFromPerct.Index).Value = .Disc_From_Perct
                        rowDGV.Cells(vFrm.ColNettPrice.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Nett_Price, RoundingItem))
                        rowDGV.Cells(vFrm.ColAmount.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(.Amount, RoundingItem))
                        rowDGV.Cells(vFrm.ColPrimaryCurrency.Index).Value = .Primary_Currency
                        rowDGV.Cells(vFrm.ColCurrency.Index).Value = vFrm.ObjCurrency.GetCurrencyCode(.Primary_Currency)
                        rowDGV.Cells(vFrm.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(ObjCD.RemoveDecimalZero(CDec(.Exchanged_Rate)))
                        rowDGV.Cells(vFrm.ColPrimaryLocation.Index).Value = .primary_location
                        rowDGV.Cells(vFrm.ColLocation.Index).Value = .Storage_Area_Code
                        rowDGV.Cells(vFrm.ColNo.Index).Value = vDGVItem.Rows.Count
                        rowDGV.Cells(vFrm.ColPurchasInIDR.Index).Value = .Purchase_Price_In_IDR

                        ''@rosadinnn
                        ''@MUaT
                        rowDGV.Cells(vFrm.ColRatioReference.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, objProduct.GetAngkaDecimal(.Primary_Product, .Primary_Unit_Multi))
                        rowDGV.Cells(vFrm.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(.Ratio_Default, objProduct.GetAngkaDecimal(.Primary_Product))
                        rowDGV.Cells(vFrm.ColPrimaryUnitMulti.Index).Value = .Primary_Unit_Multi
                        If .Primary_Units <> .Primary_Unit_Multi Then
                            rowDGV.Cells(vFrm.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, _
                                    objProduct.GetAngkaDecimal(.Primary_Product, .Primary_Unit_Multi)) _
                                    & " " & .Unit_Code_Multi & " = " & ObjCD.DigitGrouping(.Ratio_Default, _
                                    objProduct.GetAngkaDecimal(.Primary_Product)) & " " & .Unit_Code
                        Else
                            rowDGV.Cells(vFrm.ColRatioDesc.Index).Value = "-"
                        End If
                        rowDGV.Cells(vFrm.ColTotalQty.Index).Value = ObjCD.DigitGrouping(.Total_Qty, objProduct.GetAngkaDecimal(.Primary_Product))
                        rowDGV.Cells(vFrm.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, objProduct.GetAngkaDecimal(.Primary_Product, .Primary_Unit_Multi))
                        rowDGV.Cells(vFrm.ColUnitsCodeMulti.Index).Value = .Unit_Code_Multi

                        rowDGV.Cells(vFrm.colPriceByMultiUnit.Index).Value = .price_by_multi_unit
                        If .price_by_multi_unit = 1 Then
                            rowDGV.Cells(vFrm.colPriceByMultiUnit_Str.Index).Value = "/" + .Unit_Code_Multi
                        Else
                            rowDGV.Cells(vFrm.colPriceByMultiUnit_Str.Index).Value = "/" + .Unit_Code
                        End If
                    End With
                Next
            End If

            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

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

    Public Function Simpan(ByVal vdgvItem As DataGridView, _
    ByVal vPrimaryTransaksi As Double, ByVal vFrm As FrmTransaksiPurchaseReturn, _
     Optional ByVal vSetPOB As Boolean = True) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item1TableAdapter
        Dim objProduct As New Core_Module.clsMasterProduct
        Dim objitem As New ClsTransaksiPOItem
        Dim Expirydate As Date
        Dim SecQty As Double = -1
        Dim vListPrimaryPO As New List(Of Double)
        Dim vObjPOItem As New ClsTransaksiPOItem
        Dim ObjCustom As New CustomSMESS.ClsTransaksiPurchaseReturnQuick_AddByCustomProduct

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            For Each row In vdgvItem.Rows
                'simpan ke basis data item
                If IsDate(row.Cells("ColExpiryDate").Value) = True Then
                    Expirydate = row.Cells("ColExpiryDate").Value
                Else
                    Expirydate = #1/1/1950#
                End If
                If IsNumeric(row.Cells("ColSecQty").Value) = True Then
                    SecQty = row.Cells("ColSecQty").Value
                Else
                    SecQty = -1
                End If

                DTA.InsertQuery(vPrimaryTransaksi, row.Cells("ColPrimaryCodeNo").Value, row.Cells("ColPrimaryPO").Value, _
                row.Cells("ColBatchNo").Value, Expirydate, row.Cells("ColQtyMaster").Value, row.Cells("ColQtyPerMaster").Value, _
                row.Cells("ColPrimaryMasterUnits").Value, row.Cells("ColQty").Value, _
                row.Cells("ColTotalQty").Value, row.Cells("ColPrimaryUnit").Value, SecQty, _
                row.Cells("ColPrimarySecUnit").Value, row.Cells("colRemarks").Value, row.Cells(vFrm.ColUnitPrice.Index).Value, _
                row.Cells(vFrm.ColDiscount.Index).Value, row.Cells(vFrm.ColNettPrice.Index).Value, row.Cells(vFrm.ColAmount.Index).Value, _
                row.Cells(vFrm.ColExchangeRate.Index).Value, row.Cells(vFrm.ColPrimaryCurrency.Index).Value, -1, row.Cells(vFrm.ColPurchaseInIDR.Index).Value, _
                0, 1, row.Cells(vFrm.ColQtyMulti.Index).Value, row.Cells(vFrm.ColPrimaryUnitMulti.Index).Value, row.Cells(vFrm.ColRatioReference.Index).Value, _
                row.Cells(vFrm.ColRatioDefault.Index).Value, row.Cells(vFrm.ColPriceByMultiUnit.Index).Value)

                If vListPrimaryPO.Contains(row.Cells("ColPrimaryPO").Value) = False Then
                    vListPrimaryPO.Add(row.Cells("ColPrimaryPO").Value)
                End If
            Next

            If vSetPOB = True Then
                vObjPOItem.SetOrderBalance(vListPrimaryPO)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            objProduct = Nothing
            objitem = Nothing
            vListPrimaryPO = Nothing
            vObjPOItem = Nothing
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItem As DataGridView, _
    ByVal vPrimaryTransaksi As Double, ByVal vFrm As FrmTransaksiPurchaseReturn_Quick, _
       ByVal vTblCustom As CustomSMESS.DseTransaksiPurchaseReturnQuick_AddByCustomProduct.tbltrans_custom_PR_item1DataTableDataTable, _
    Optional ByVal vSetPOB As Boolean = True) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_itemTableAdapter
        Dim objProduct As New Core_Module.clsMasterProduct
        Dim objitem As New ClsTransaksiPOItem
        Dim Expirydate As Date
        Dim SecQty As Double = -1
        Dim vListPrimaryPO As New List(Of Double)
        Dim vObjPOItem As New ClsTransaksiPOItem
        Dim ObjCustom As New CustomSMESS.ClsTransaksiPurchaseReturnQuick_AddByCustomProduct

        Try
            For Each row In vdgvItem.Rows
                'simpan ke basis data item
                If IsDate(row.Cells("ColExpiryDate").Value) = True Then
                    Expirydate = row.Cells("ColExpiryDate").Value
                Else
                    Expirydate = #1/1/1950#
                End If
                If IsNumeric(row.Cells("ColSecQty").Value) = True Then
                    SecQty = row.Cells("ColSecQty").Value
                Else
                    SecQty = -1
                End If

                DTA.Insert_Multi(vPrimaryTransaksi, row.Cells("ColPrimaryCodeNo").Value, row.Cells("ColPrimaryPO").Value, _
                row.Cells("ColBatchNo").Value, _
                Expirydate, row.Cells("ColQtyMaster").Value, row.Cells("ColQtyPerMaster").Value, _
                row.Cells("ColPrimaryMasterUnits").Value, row.Cells("ColQty").Value, _
                row.Cells("ColTotalQty").Value, row.Cells("ColPrimaryUnit").Value, SecQty, _
                row.Cells("ColPrimarySecUnit").Value, row.Cells("colRemarks").Value, row.Cells(vFrm.ColUnitPrice.Index).Value, _
                row.Cells(vFrm.ColDiscount.Index).Value, row.Cells(vFrm.ColNettPrice.Index).Value, row.Cells(vFrm.ColAmount.Index).Value, _
                row.Cells(vFrm.ColExchangeRate.Index).Value, row.Cells(vFrm.ColPrimaryCurrency.Index).Value, row.Cells(vFrm.ColPrimaryLocation.Index).Value, _
                row.Cells(vFrm.ColPurchasInIDR.Index).Value, _
                row.Cells(vFrm.ColTotalQtyMulti.Index).Value, row.Cells(vFrm.ColPrimaryUnitMulti.Index).Value, _
                row.Cells(vFrm.ColRatioReference.Index).Value, row.Cells(vFrm.ColRatioDefault.Index).Value, _
                row.Cells(vFrm.ColDiscAmount.Index).Value, row.Cells(vFrm.ColDiscFromPerct.Index).Value, row.Cells(vFrm.colPriceByMultiUnit.Index).Value)

                If vListPrimaryPO.Contains(row.Cells("ColPrimaryPO").Value) = False Then
                    vListPrimaryPO.Add(row.Cells("ColPrimaryPO").Value)
                End If
            Next

            If vSetPOB = True Then
                vObjPOItem.SetOrderBalance(vListPrimaryPO)
            End If
            '//TR
            ObjCustom.SimpanFromTblDummy(vPrimaryTransaksi, vTblCustom)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            objProduct = Nothing
            objitem = Nothing
            vListPrimaryPO = Nothing
            vObjPOItem = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vObjPR As ClsTransaksiPurchaseReturn) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_itemTableAdapter
        'Dim DTA2 As New CustomSMESS.DseTransaksiPurchaseReturnQuick_AddByCustomProductTableAdapters.tbltrans_custom_purchasereturn_itemTableAdapter
        Dim vListPrimaryPO As New List(Of Double)
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim ObjCustom As New CustomSMESS.ClsTransaksiPurchaseReturnQuick_AddByCustomProduct
        Try
            vObjPR.GetListPrimaryPO(vPrimaryTransaksi, vListPrimaryPO)
            'hapus data item
            DTA.Delete(vPrimaryTransaksi)
            '//TR
            ObjCustom.Hapus(vPrimaryTransaksi)

            'DTA2.Delete(vPrimaryTransaksi)
            'set order balance
            ObjPOItem.SetOrderBalance(vListPrimaryPO)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjPOItem = Nothing
            vListPrimaryPO = Nothing
            ObjCustom = Nothing
        End Try
    End Function
End Class
