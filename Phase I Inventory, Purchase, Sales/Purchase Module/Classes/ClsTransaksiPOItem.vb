Option Explicit On
Option Strict Off

Public Class ClsTransaksiPOItem
    Private mPrimary As Double = -1
    Private mprimaryTransaksi As Double = -1
    Private mprimaryproduct As Double = -1
    Private mprimaryIPO As String = "" ''Sementara
    Private mQty As Double = 0
    Private mPrimaryUnits As Double = -1
    Private mDeliveryDateCont As Boolean
    Private mdeliverydate As Date = Now
    Private munitprice As Double = 0
    Private mDiscount As Double = 0
    Private mNettPrice As Double = 0
    Private mAmount As Double = 0
    Private mRemark As String
    Private mUseSecondary As Integer

    Public Property UseSecondary() As String
        Get
            Return Me.mUseSecondary
        End Get
        Set(ByVal value As String)
            mUseSecondary = value
        End Set
    End Property
    Public Property Remark() As String
        Get
            Return Me.mRemark
        End Get
        Set(ByVal value As String)
            mRemark = value
        End Set
    End Property
    Public Property Amount() As Double
        Get
            Return Me.mAmount
        End Get
        Set(ByVal value As Double)
            mAmount = value
        End Set
    End Property

    Public Property NettPrice() As Double
        Get
            Return Me.mNettPrice
        End Get
        Set(ByVal value As Double)
            mNettPrice = value
        End Set
    End Property

    Public Property Diskon() As Double
        Get
            Return Me.mDiscount
        End Get
        Set(ByVal value As Double)
            mDiscount = value
        End Set
    End Property
    Public Property unitPrice() As Double
        Get
            Return Me.munitprice
        End Get
        Set(ByVal value As Double)
            munitprice = value
        End Set
    End Property
    Public Property DeliveryDate() As Date
        Get
            Return Me.mdeliverydate
        End Get
        Set(ByVal value As Date)
            mdeliverydate = value
        End Set
    End Property

    Public Property DeliveryDateCont() As Boolean
        Get
            Return Me.mDeliveryDateCont
        End Get
        Set(ByVal value As Boolean)
            mDeliveryDateCont = value
        End Set
    End Property
    Public Property PrimaryUnits() As Double
        Get
            Return Me.mPrimaryUnits
        End Get
        Set(ByVal value As Double)
            mPrimaryUnits = value
        End Set
    End Property
    Public Property Quantity() As Double
        Get
            Return Me.mQty
        End Get
        Set(ByVal value As Double)
            mQty = value
        End Set
    End Property
    '' Sementara 
    Public Property PrimaryIPO() As String
        Get
            Return Me.mprimaryIPO
        End Get
        Set(ByVal value As String)
            mprimaryIPO = value
        End Set
    End Property

    Public Property PrimaryProduct() As Double
        Get
            Return mprimaryproduct
        End Get
        Set(ByVal value As Double)
            mprimaryproduct = value
        End Set
    End Property

    Public Property PrimaryTransaksi() As Double
        Get
            Return mprimaryTransaksi
        End Get
        Set(ByVal value As Double)
            mprimaryTransaksi = value
        End Set
    End Property

    Public Property primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Function GetNoIPO(ByVal VprimaryIPO As Double) As String
        Dim dta As New DseTransaksiPOTableAdapters.Tbl_IPO_1TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(dta.ScalarQuery_GetNoIPO(VprimaryIPO)) = False Then
                Return dta.ScalarQuery_GetNoIPO(VprimaryIPO)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal DefaultPrec As Integer, _
    ByVal vTBLItemDelivery As DseTransaksiPO.tbltrans_po_itemdeliveryDataTable, ByVal vFrm As FrmTransaksiPO, _
    ByVal vObjPO As ClsTransaksiPO) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_po_item5TableAdapter
        Dim TBL As New DseTransaksiPO.tbltrans_po_item5DataTable
        Dim row As DseTransaksiPO.tbltrans_po_item5Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim objProduct As New Core_Module.clsMasterProduct
        Dim objitem As New ClsTransaksiPOItem
        Dim objUnit As New Core_Module.ClsMasterUnit
        Dim DTADelivery As New DseTransaksiPOTableAdapters.tbltrans_po_itemdeliveryTableAdapter
        Dim DelDateStr As String = ""
        Dim rowDGV As DataGridViewRow
        Dim AdaPriceByUnitDefault As Boolean = False

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTADelivery.Connection.ConnectionString = DTA.Connection.ConnectionString

            'item delivery
            DTADelivery.FillBy_Primary(vTBLItemDelivery, vPrimaryTransaksi)

            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        If CBool(.Delivery_Date_Cont) = True Then
                            vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + row.identifier + "'"
                            If vTBLItemDelivery.DefaultView.Count = 1 Then
                                DelDateStr = Format(CDate(.Delivery_Date), "dd MMM yyyy")
                            Else
                                DelDateStr = "Multiple"
                            End If
                        Else
                            DelDateStr = "-"
                        End If

                        vDGVItem.Rows.Add()
                        rowDGV = vDGVItem.Rows(vDGVItem.RowCount - 1)
                        rowDGV.Cells(vFrm.ColNo.Index).Value = vDGVItem.RowCount
                        rowDGV.Cells(vFrm.ColCodeNo.Index).Value = .Product_Code
                        rowDGV.Cells(vFrm.ColDescription.Index).Value = .Description
                        rowDGV.Cells(vFrm.ColQty.Index).Value = ObjCD.DigitGrouping(.Qty, ._Decimal)
                        rowDGV.Cells(vFrm.ColUnit.Index).Value = .Unit_Code
                        rowDGV.Cells(vFrm.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(.Unit_Price, DefaultPrec)
                        rowDGV.Cells(vFrm.ColDiscount.Index).Value = ObjCD.DigitGrouping(.Discount, 2)
                        rowDGV.Cells(vFrm.ColDiscAmount.Index).Value = ObjCD.DigitGrouping(.Disc_Amount, DefaultPrec)
                        rowDGV.Cells(vFrm.ColDiscFromPerct.Index).Value = .Disc_From_Perct
                        rowDGV.Cells(vFrm.ColNettPrice.Index).Value = ObjCD.DigitGrouping(.Nett_Price, DefaultPrec)
                        rowDGV.Cells(vFrm.ColAmount.Index).Value = ObjCD.DigitGrouping(.Amount, DefaultPrec)
                        rowDGV.Cells(vFrm.ColDeliveryCont.Index).Value = CBool(.Delivery_Date_Cont)
                        rowDGV.Cells(vFrm.ColDeliveryDate.Index).Value = DelDateStr
                        rowDGV.Cells(vFrm.ColRemark.Index).Value = .Remarks
                        rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rowDGV.Cells(vFrm.ColPrimaryUnits.Index).Value = .Primary_Units
                        rowDGV.Cells(vFrm.ColIPONo.Index).Value = ""
                        rowDGV.Cells(vFrm.ColPrimaryIPO.Index).Value = .Primary_IPO
                        rowDGV.Cells(vFrm.colUserSecondaryUnits.Index).Value = CBool(.Use_Secondary)
                        rowDGV.Cells(vFrm.ColIdentifier.Index).Value = .identifier

                        ''@rosadinnn
                        ''@MUaT
                        rowDGV.Cells(vFrm.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(.Qty_Multi, objProduct.GetAngkaDecimal(.Primary_Product, .Primary_Unit_Multi))
                        rowDGV.Cells(vFrm.ColPrimaryUnitMulti.Index).Value = .Primary_Unit_Multi.ToString
                        rowDGV.Cells(vFrm.ColUnitCodeMulti.Index).Value = objUnit.GetUnitCode(.Primary_Unit_Multi)
                        rowDGV.Cells(vFrm.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(.Ratio_Default, objProduct.GetAngkaDecimal(rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value))
                        rowDGV.Cells(vFrm.ColRatioReference.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, objProduct.GetAngkaDecimal(rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value, rowDGV.Cells(vFrm.ColPrimaryUnitMulti.Index).Value))
                        vFrm.SetRatioAndQtyDefault(rowDGV, False)
                        rowDGV.Cells(vFrm.ColToleranceDown.Index).Value = ObjCD.DigitGrouping(.Tolerance_Down, objProduct.GetAngkaDecimal(rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value, rowDGV.Cells(vFrm.ColPrimaryUnitMulti.Index).Value))
                        rowDGV.Cells(vFrm.ColToleranceUp.Index).Value = ObjCD.DigitGrouping(.Tolerance_Up, objProduct.GetAngkaDecimal(rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value, rowDGV.Cells(vFrm.ColPrimaryUnitMulti.Index).Value))
                        rowDGV.Cells(vFrm.ColToleranceDownPerc.Index).Value = ObjCD.DigitGrouping(.Tolerance_Down / .Qty_Multi * 100, 2)
                        rowDGV.Cells(vFrm.ColToleranceUpPerc.Index).Value = ObjCD.DigitGrouping(.Tolerance_Up / .Qty_Multi * 100, 2)

                        rowDGV.Cells(vFrm.ColPriceByMultiUnit.Index).Value = .price_by_multi_unit
                        If .price_by_multi_unit = 1 Then
                            rowDGV.Cells(vFrm.ColPriceByMultiUnit_Str.Index).Value = "/" + rowDGV.Cells(vFrm.ColUnitCodeMulti.Index).Value
                        Else
                            rowDGV.Cells(vFrm.ColPriceByMultiUnit_Str.Index).Value = "/" + .Unit_Code
                        End If
                        If .price_by_multi_unit = 0 Then
                            AdaPriceByUnitDefault = True
                        End If
                    End With
                Next
            End If

            'set nomer
            ObjCD.SetNumber(vDGVItem)

            vFrm.ColPriceByMultiUnit_Str.Visible = AdaPriceByUnitDefault
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            objProduct = Nothing
            objitem = Nothing
            objUnit = Nothing
            DTADelivery.Dispose()
        End Try
    End Function

    Public Function GetJmlQty(ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double) As Decimal
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_po_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJmlQtyByPO_n_Product(vPrimaryProduct, vPrimaryPO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryIPO(ByVal VnoIPO As String) As Double
        Dim dta As New DseTransaksiPOTableAdapters.Tbl_IPO_1TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.ScalarQuery_GetPrimaryIPO(VnoIPO).HasValue Then
                Return dta.ScalarQuery_GetPrimaryIPO(VnoIPO).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItem As DataGridView, _
    ByVal vPrimaryTransaksi As Double, ByVal vFrm As FrmTransaksiPO, ByVal vTBLDelivery As DseTransaksiPO.tbltrans_po_itemdeliveryDataTable) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_po_itemTableAdapter
        Dim DelDateCont, UseSecondary As Byte
        Dim DelDate As Date
        Dim PriceCondition As Integer = 0
        Dim objProduct As New Core_Module.clsMasterProduct
        Dim objitem As New ClsTransaksiPOItem
        Dim rowDelivery As DseTransaksiPO.tbltrans_po_itemdeliveryRow
        Dim DTADelivery As New DseTransaksiPOTableAdapters.tbltrans_po_itemdeliveryTableAdapter
        Dim Int As Integer = 0
        Dim DTAUpdateProduct As New Sales_Module.DseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter
        Dim listPrimaryProduct As New List(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTADelivery.Connection.ConnectionString = DTA.Connection.ConnectionString

            For Each row In vdgvItem.Rows
                'simpan ke basis data item
                If row.Cells("ColDeliveryCont").Value = Boolean.TrueString Then
                    DelDateCont = 1
                    vTBLDelivery.DefaultView.RowFilter = "identifier = '" + row.Cells("ColIdentifier").Value + "'"
                    If vTBLDelivery.DefaultView.Count = 1 Then
                        DelDate = CDate(row.Cells("ColDeliveryDate").Value)
                    Else
                        DelDate = #1/1/1950#
                    End If
                Else
                    DelDateCont = 0
                    DelDate = #1/1/1950#
                End If
                If row.Cells("colUserSecondaryUnits").Value = Boolean.TrueString Then
                    UseSecondary = 1
                Else
                    UseSecondary = 0
                End If

                If row.Cells("ColIdentifier").Value.ToString.StartsWith("-1") Then
                    row.Cells("ColIdentifier").Value = vPrimaryTransaksi.ToString + row.Cells("ColIdentifier").Value.ToString.Substring(2)
                End If

                ''@rosadinnn
                ''@MUaT
                ''@update insert query
                DTA.Insert(vPrimaryTransaksi, row.Cells("ColPrimaryCodeNo").Value, row.Cells("ColPrimaryIPO").Value, row.Cells("ColQty").Value, _
                row.Cells("ColPrimaryUnits").Value, DelDateCont, DelDate, row.Cells("ColUnitPrice").Value, row.Cells("ColDiscount").Value, _
                row.Cells("ColNettPrice").Value, row.Cells("ColAmount").Value, row.Cells("ColRemark").Value, UseSecondary, row.Cells("ColIdentifier").Value, _
                row.Cells(vFrm.ColDiscAmount.Index).Value, row.Cells(vFrm.ColDiscFromPerct.Index).Value, _
                row.Cells(vFrm.ColRatioDefault.Index).Value, row.Cells(vFrm.ColRatioReference.Index).Value, _
                row.Cells(vFrm.ColQtyMulti.Index).Value, row.Cells(vFrm.ColPrimaryUnitMulti.Index).Value, _
                row.Cells(vFrm.ColToleranceDown.Index).Value, row.Cells(vFrm.ColToleranceUp.Index).Value, row.Cells(vFrm.ColPriceByMultiUnit.Index).Value)

                If listPrimaryProduct.Contains(row.Cells(vFrm.ColPrimaryCodeNo.Index).Value) = False Then
                    listPrimaryProduct.Add(row.Cells(vFrm.ColPrimaryCodeNo.Index).Value)
                    DTAUpdateProduct.UpdateQuery_SetTblMasterProduct(row.Cells(vFrm.ColPriceByMultiUnit.Index).Value, row.Cells(vFrm.ColPrimaryCodeNo.Index).Value)
                End If
            Next

            For Each rowDelivery In vTBLDelivery.Rows
                With rowDelivery
                    If .identifier.StartsWith("-1") = True Then
                        .identifier = vPrimaryTransaksi.ToString + .identifier.Substring(2)
                    End If

                    ''@rosadinnn
                    ''@MUaT
                    ''simpan delivery qty multi
                    DTADelivery.Insert(vPrimaryTransaksi, .identifier, .delivery_qty, .delivery_date, .remarks, .delay, .delivery_qty_multi)

                End With
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            objProduct = Nothing
            objitem = Nothing
            DTADelivery.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_po_itemTableAdapter
        Dim DTA1 As New DseTransaksiPOTableAdapters.tbltrans_po_itemdeliveryTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            'hapus data item
            DTA.Delete(vPrimaryTransaksi)
            DTA1.Delete(vPrimaryTransaksi)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
        End Try
    End Function

    Public Sub IsikanTBLProduct_Default(ByVal vPrimaryPO As Double, _
    ByVal TBL As DseTransaksiPO.orderbalance_report_view0DataTable, ByVal vCBo As ComboBox)
        Dim DTA As New DseTransaksiPOTableAdapters.orderbalance_report_view0TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryPO(TBL, vPrimaryPO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Public Sub IsikanTBLProduct_Multi(ByVal vPrimaryPO As Double, _
    ByVal TBL As DseTransaksiPO.orderbalance_report_view0MultiDataTable, ByVal vCBo As ComboBox, ByVal vIgnoreTolerance As Boolean)
        Dim DTA As New DseTransaksiPOTableAdapters.orderbalance_report_view0MultiTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If vIgnoreTolerance = False Then
                DTA.FillBy_PrimaryPO_Tolerance(TBL, vPrimaryPO)
            Else
                DTA.FillBy_PrimaryPO(TBL, vPrimaryPO)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function IsikanTBLPO_Default(ByVal vPrimarySupplier As Double, ByVal TBL As DseTransaksiPO.orderbalance_report_view0DataTable, _
    ByVal cbo As ComboBox, Optional ByVal vNonQuickPOOnly As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.orderbalance_report_view0TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If vNonQuickPOOnly = True Then
                DTA.FillBy_PrimarySupplier2(TBL, vPrimarySupplier)
            Else
                DTA.FillBy_PrimarySupplier(TBL, vPrimarySupplier)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function IsikanTBLPO_Multi(ByVal vPrimarySupplier As Double, ByVal TBL As DseTransaksiPO.orderbalance_report_view0MultiDataTable, _
    ByVal cbo As ComboBox, Optional ByVal vNonQuickPOOnly As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.orderbalance_report_view0MultiTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If vNonQuickPOOnly = True Then
                DTA.FillBy_PrimarySupplier2(TBL, vPrimarySupplier)
            Else
                DTA.FillBy_PrimarySupplier(TBL, vPrimarySupplier)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsikanTBLPO_Multi_Tolerance(ByVal vPrimarySupplier As Double, ByVal TBL As DseTransaksiPO.orderbalance_report_view0MultiDataTable, _
    ByVal cbo As ComboBox, Optional ByVal vNonQuickPOOnly As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.orderbalance_report_view0MultiTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If vNonQuickPOOnly = True Then
                DTA.FillBy_PrimarySupplier2_Tolerance(TBL, vPrimarySupplier)
            Else
                DTA.FillBy_PrimarySupplier_Tolerance(TBL, vPrimarySupplier)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function GetJmlPR(ByVal vPrimaryProduct As Double, _
    ByVal vTBLPR As DseTransaksiPO_SetOrderBalance.tbltrans_purchasereturn_itemDataTable, _
    ByVal vQtyMulti As Boolean) As Decimal
        Dim Filter As String = "[Primary Product] = " + vPrimaryProduct.ToString + ""
        Dim Dummy As Double = 0
        Dim rowPR As DseTransaksiPO_SetOrderBalance.tbltrans_purchasereturn_itemRow

        Try
            vTBLPR.DefaultView.RowFilter = Filter
            If vTBLPR.DefaultView.Count > 0 Then
                rowPR = vTBLPR.DefaultView(0).Row

                ''@rosadinnn
                ''@MUaT
                If vQtyMulti = True Then
                    Return rowPR.Qty_Multi
                End If
                Return rowPR.Qty

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJmlPurchaseOrderBalance_Default(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double) As Decimal
        Dim DTA As New DseTransaksiPOTableAdapters.orderbalance_report_view0TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlOrderBalance(vPrimaryPO, vPrimaryProduct).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlOrderBalance(vPrimaryPO, vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function GetJmlPurchaseOrderBalance_Multi(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double) As Decimal
        Dim DTA As New DseTransaksiPOTableAdapters.orderbalance_report_view0MultiTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlOrderBalanceMulti(vPrimaryPO, vPrimaryProduct).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlOrderBalanceMulti(vPrimaryPO, vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlPOC(ByVal vPrimaryProduct As Double, _
    ByVal vTBLPOC As DseTransaksiPO_SetOrderBalance.tbltrans_poc_itemDataTable, _
    ByVal vQtyMulti As Boolean) As Decimal
        Dim Filter As String = "[Primary Product] = " + vPrimaryProduct.ToString + ""
        Dim Dummy As Double = 0
        Dim rowPOC As DseTransaksiPO_SetOrderBalance.tbltrans_poc_itemRow

        Try
            vTBLPOC.DefaultView.RowFilter = Filter
            If vTBLPOC.DefaultView.Count > 0 Then
                rowPOC = vTBLPOC.DefaultView(0).Row

                ''@rosadinnn
                ''@MUaT
                If vQtyMulti = True Then
                    Return rowPOC.Qty_Multi
                End If
                Return rowPOC.Qty

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJmlGRN(ByVal vPrimaryProduct As Double, _
    ByVal vTBLGRN As DseTransaksiPO_SetOrderBalance.tbltrans_grn_itemDataTable, _
    ByVal vQtyMulti As Boolean) As Decimal
        Dim Filter As String = "[Primary Product] = " + vPrimaryProduct.ToString + ""
        Dim Dummy As Double = 0
        Dim rowGRN As DseTransaksiPO_SetOrderBalance.tbltrans_grn_itemRow

        Try
            vTBLGRN.DefaultView.RowFilter = Filter
            If vTBLGRN.DefaultView.Count > 0 Then
                rowGRN = vTBLGRN.DefaultView(0).Row

                ''@rosadinnn
                ''@MUaT
                If vQtyMulti = True Then
                    Return rowGRN.Qty_Multi
                End If

                Return rowGRN.Qty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub GetHargaDetail(ByVal vPrimaryPO As Double, ByVal vPrimaryCodeNo As Double, ByRef mUnitPrice As Decimal, ByRef mDisc As Decimal, _
    ByRef mNettPrice As Decimal)
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_po_itemTableAdapter
        Dim TBL As New DseTransaksiPO.tbltrans_po_itemDataTable
        Dim row As DseTransaksiPO.tbltrans_po_itemRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_PrimaryTransaksi_PrimaryProduct(TBL, vPrimaryPO, vPrimaryCodeNo)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)

                mUnitPrice = row.Unit_Price
                mDisc = row.Discount
                mNettPrice = row.Nett_Price
            Else
                mUnitPrice = 0
                mDisc = 0
                mNettPrice = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetJmlGIT(ByVal vPrimaryProduct As Double, _
    ByVal vTBLGIT As DseTransaksiPO_SetOrderBalance.tbltrans_packinglist_itemDataTable, _
    ByVal vQtyMulti As Boolean) As Decimal
        Dim Filter As String = "[Primary Product] = " + vPrimaryProduct.ToString + ""
        Dim Dummy As Double = 0
        Dim rowGIT As DseTransaksiPO_SetOrderBalance.tbltrans_packinglist_itemRow

        Try
            vTBLGIT.DefaultView.RowFilter = Filter
            If vTBLGIT.DefaultView.Count > 0 Then
                rowGIT = vTBLGIT.DefaultView(0).Row

                ''@rosadinnn
                ''@MUaT
                If vQtyMulti = True Then
                    Return rowGIT.Qty_Multi
                End If
                Return rowGIT.Total_Qty

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    ''@rosadinnn
    ''@MUaT
    ''@tambahkan kolom pob multi
    Public Sub SetOrderBalance(ByVal vPrimaryPO As Double)
        Dim TBLPO As New DseTransaksiPO_SetOrderBalance.tbltrans_po_itemDataTable
        Dim DTAPO As New DseTransaksiPO_SetOrderBalanceTableAdapters.tbltrans_po_itemTableAdapter
        Dim rowPO As DseTransaksiPO_SetOrderBalance.tbltrans_po_itemRow
        Dim TBLPR As New DseTransaksiPO_SetOrderBalance.tbltrans_purchasereturn_itemDataTable
        Dim DTAPR As New DseTransaksiPO_SetOrderBalanceTableAdapters.tbltrans_purchasereturn_itemTableAdapter
        Dim TBLPOC As New DseTransaksiPO_SetOrderBalance.tbltrans_poc_itemDataTable
        Dim DTAPOC As New DseTransaksiPO_SetOrderBalanceTableAdapters.tbltrans_poc_itemTableAdapter
        Dim TBLGIT As New DseTransaksiPO_SetOrderBalance.tbltrans_packinglist_itemDataTable
        Dim DTAGIT As New DseTransaksiPO_SetOrderBalanceTableAdapters.tbltrans_packinglist_itemTableAdapter
        Dim TBLGRN As New DseTransaksiPO_SetOrderBalance.tbltrans_grn_itemDataTable
        Dim DTAGRN As New DseTransaksiPO_SetOrderBalanceTableAdapters.tbltrans_grn_itemTableAdapter
        Dim JmlPO, JmlPR, JmlGIT, JmlGRN, JmlPOC, JMLPOB As Decimal
        Dim JmlTotalGIT, JmlTotalGRN, JmlTotalDP As Decimal

        ''@rosadinnn
        ''@MUaT
        Dim JmlPO_Multi, JmlPR_Multi, JmlGIT_Multi, JmlGRN_Multi, JmlPOC_Multi, JMLPOB_Multi As Decimal
        Dim dtaRatio As New DseTransaksiPOTableAdapters.tbltrans_po_item6TableAdapter
        Dim tblRatio As New DseTransaksiPO.tbltrans_po_item6DataTable
        Dim rowRatio As DseTransaksiPO.tbltrans_po_item6Row
        Dim ratioRef As Decimal = 1
        Dim ratioDef As Decimal = 1

        Dim TBLDP As New DseTransaksiAOP.tbltrans_aopDataTable
        Dim DTADP As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter

        Try
            DTAPO.Connection.ConnectionString = DTAPO.Connection.ConnectionString
            DTAPR.Connection.ConnectionString = DTAPO.Connection.ConnectionString
            DTAPOC.Connection.ConnectionString = DTAPO.Connection.ConnectionString
            DTAGIT.Connection.ConnectionString = DTAPO.Connection.ConnectionString
            DTAGRN.Connection.ConnectionString = DTAPO.Connection.ConnectionString
            dtaRatio.Connection.ConnectionString = DTAPO.Connection.ConnectionString
            DTADP.Connection.ConnectionString = DTAPO.Connection.ConnectionString

            'fill data table
            JmlTotalGIT = 0
            JmlTotalGRN = 0
            JmlTotalDP = 0

            DTAPO.FillBy_PrimaryPO(TBLPO, vPrimaryPO)
            DTAPR.FillBy_PrimaryPO(TBLPR, vPrimaryPO)
            DTAPOC.FillBy_PrimaryPO(TBLPOC, vPrimaryPO)
            DTAGIT.FillBy_PrimaryPO(TBLGIT, vPrimaryPO)
            DTAGRN.FillBy_PrimaryPO(TBLGRN, vPrimaryPO)

            ''@rosadinnn
            ''@MUaT
            dtaRatio.Fill(tblRatio, vPrimaryPO)

            For Each rowPO In TBLPO.Rows
                JmlPO = rowPO.Qty
                JmlPR = Me.GetJmlPR(rowPO.Primary_Product, TBLPR, False)
                JmlPOC = Me.GetJmlPOC(rowPO.Primary_Product, TBLPOC, False)
                JmlGRN = Me.GetJmlGRN(rowPO.Primary_Product, TBLGRN, False)
                JmlGIT = Me.GetJmlGIT(rowPO.Primary_Product, TBLGIT, False)
                JMLPOB = JmlPO + JmlPR - JmlGIT - JmlPOC - JmlGRN

                ''@rosadinnn
                ''@MUaT
                JmlPO_Multi = rowPO.Qty_Multi
                JmlPR_Multi = Me.GetJmlPR(rowPO.Primary_Product, TBLPR, True)
                JmlPOC_Multi = Me.GetJmlPOC(rowPO.Primary_Product, TBLPOC, True)
                JmlGRN_Multi = Me.GetJmlGRN(rowPO.Primary_Product, TBLGRN, True)
                JmlGIT_Multi = Me.GetJmlGIT(rowPO.Primary_Product, TBLGIT, True)

                ''ubah cara perhitungan POB default(dg menerjemahkan POB multi)
                tblRatio.DefaultView.RowFilter = "[Primary Product] = " & rowPO.Primary_Product.ToString
                If tblRatio.DefaultView.Count = 0 Then
                    MsgBox("Logic error, while trying set POB default(primary product: " & rowPO.Primary_Product.ToString & "). Please contact your software provider.", MsgBoxStyle.Information, "Information")
                    ratioRef = 1
                    ratioDef = 1
                Else
                    rowRatio = tblRatio.DefaultView(0).Row
                    ratioRef = rowRatio.Ratio_Reference
                    ratioDef = rowRatio.Ratio_Default
                End If
                JMLPOB_Multi = JmlPO_Multi + JmlPR_Multi - JmlGIT_Multi - JmlPOC_Multi - JmlGRN_Multi
                JMLPOB = Math.Round(JMLPOB_Multi * ratioDef / ratioRef, rowPO._Decimal, MidpointRounding.AwayFromZero)

                DTAPO.UpdateQuery_setGIT(JmlGIT, JmlGIT_Multi, rowPO.Primary_PO, rowPO.Primary_Product)
                DTAPO.UpdateQuery_SetPOB(JMLPOB, JMLPOB_Multi, rowPO.Primary_PO, rowPO.Primary_Product)

                JmlTotalGIT += JmlGIT + JmlPOC
                JmlTotalGRN += JmlGRN
            Next

            DTADP.FillBy_PrimaryPO_DP(TBLDP, vPrimaryPO)
            JmlTotalDP = TBLDP.Rows.Count

            If JmlTotalGRN > 0 Or JmlTotalGIT > 0 Or JmlTotalDP > 0 Then
                DTAPO.UpdateQuery(JmlTotalGRN + JmlTotalGIT + JmlTotalDP, vPrimaryPO)
            Else
                DTAPO.UpdateQuery(0, vPrimaryPO)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLPO.Dispose()
            DTAPO.Dispose()
            TBLPR.Dispose()
            DTAPR.Dispose()
            TBLPOC.Dispose()
            DTAPOC.Dispose()
            TBLGIT.Dispose()
            DTAGIT.Dispose()
            TBLGRN.Dispose()
            DTAGRN.Dispose()
            dtaRatio.Dispose()
            DTADP.Dispose()
        End Try
    End Sub

    Public Sub SetOrderBalance(ByVal vListPrimaryPO As List(Of Double))
        Dim PrimaryPO As Double

        Try
            For Each PrimaryPO In vListPrimaryPO
                If PrimaryPO <> -1 Then
                    Me.SetOrderBalanceNew(PrimaryPO)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetPreparationDelay(ByVal vPrimaryProduct As Double) As Integer
        Dim DTA1 As New DseTransaksiPOTableAdapters.tbltrans_po_itemdeliveryTableAdapter
        Dim DTA2 As New DseTransaksiPOTableAdapters.tbltrans_po_itemdelivery1TableAdapter
        Dim TBL2 As New DseTransaksiPO.tbltrans_po_itemdelivery1DataTable
        Dim row2 As DseTransaksiPO.tbltrans_po_itemdelivery1Row

        Dim PrimaryTrans As Nullable(Of Double)
        Dim RetVal As String = ""

        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = DTA1.Connection.ConnectionString

            PrimaryTrans = DTA1.ScalarQuery_GetPrimaryTransaksi(vPrimaryProduct)
            If PrimaryTrans.HasValue = True Then
                DTA2.Fill(TBL2, PrimaryTrans.Value, vPrimaryProduct)
                If TBL2.Rows.Count > 0 Then
                    'berdasarkan data delay untuk pengiriman terakhir
                    row2 = TBL2.Rows(0)
                    Return row2.delay
                End If
            End If

            'ambil berdasarkan nilai default dari setting
            Dim ObjSetting As New Core_Module.ClsTransaksiSetting
            RetVal = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PO, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PreparationDelay)
            If RetVal = "" Then
                Return 0
            Else
                Return CInt(RetVal)
            End If
            ObjSetting = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA1.Dispose()
            DTA2.Dispose()
            TBL2.Dispose()
        End Try
    End Function

    Private Function GetJmlPRNewVer(ByVal vPrimaryProduct As Double, _
   ByVal vTBLPR As DseTransaksiPO_SetOrderBalance.tbltrans_purchasereturn_itemDataTable, _
   ByVal vQtyMulti As Boolean) As Decimal

        Dim rowPR As DseTransaksiPO_SetOrderBalance.tbltrans_purchasereturn_itemRow

        Try
            ' Loop manual (lebih cepat dari RowFilter di banyak kasus)
            For Each rowPR In vTBLPR.Rows

                If rowPR.Primary_Product = vPrimaryProduct Then

                    If vQtyMulti = True Then
                        Return rowPR.Qty_Multi
                    Else
                        Return rowPR.Qty
                    End If

                End If

            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

        Return 0D

    End Function

    Public Function GetJmlPOCNewVer(ByVal vPrimaryProduct As Double, _
    ByVal vTBLPOC As DseTransaksiPO_SetOrderBalance.tbltrans_poc_itemDataTable, _
    ByVal vQtyMulti As Boolean) As Decimal

        Dim rowPOC As DseTransaksiPO_SetOrderBalance.tbltrans_poc_itemRow

        Try
            ' Loop manual (lebih ringan dari RowFilter)
            For Each rowPOC In vTBLPOC.Rows

                If rowPOC.Primary_Product = vPrimaryProduct Then

                    If vQtyMulti = True Then
                        Return rowPOC.Qty_Multi
                    Else
                        Return rowPOC.Qty
                    End If

                End If

            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

        Return 0D

    End Function

    Public Function GetJmlGRNNewVer(ByVal vPrimaryProduct As Double, _
    ByVal vTBLGRN As DseTransaksiPO_SetOrderBalance.tbltrans_grn_itemDataTable, _
    ByVal vQtyMulti As Boolean) As Decimal

        Dim rowGRN As DseTransaksiPO_SetOrderBalance.tbltrans_grn_itemRow

        Try
            ' Loop manual (lebih cepat dari RowFilter)
            For Each rowGRN In vTBLGRN.Rows

                If rowGRN.Primary_Product = vPrimaryProduct Then

                    If vQtyMulti = True Then
                        Return rowGRN.Qty_Multi
                    Else
                        Return rowGRN.Qty
                    End If

                End If

            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

        Return 0D

    End Function

    Public Function GetJmlGITNewVer(ByVal vPrimaryProduct As Double, _
    ByVal vTBLGIT As DseTransaksiPO_SetOrderBalance.tbltrans_packinglist_itemDataTable, _
    ByVal vQtyMulti As Boolean) As Decimal

        Dim rowGIT As DseTransaksiPO_SetOrderBalance.tbltrans_packinglist_itemRow

        Try
            ' Loop manual (lebih cepat dari RowFilter)
            For Each rowGIT In vTBLGIT.Rows

                If rowGIT.Primary_Product = vPrimaryProduct Then

                    If vQtyMulti = True Then
                        Return rowGIT.Qty_Multi
                    Else
                        Return rowGIT.Total_Qty   ' <-- tetap pakai ini!
                    End If

                End If

            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

        Return 0D

    End Function

    ''@rosadinnn
    ''@MUaT
    ''@tambahkan kolom pob multi
    Public Sub SetOrderBalanceNew(ByVal vPrimaryPO As Double)
        Dim TBLPO As New DseTransaksiPO_SetOrderBalance.tbltrans_po_itemDataTable
        Dim DTAPO As New DseTransaksiPO_SetOrderBalanceTableAdapters.tbltrans_po_itemTableAdapter
        Dim rowPO As DseTransaksiPO_SetOrderBalance.tbltrans_po_itemRow
        Dim TBLPR As New DseTransaksiPO_SetOrderBalance.tbltrans_purchasereturn_itemDataTable
        Dim DTAPR As New DseTransaksiPO_SetOrderBalanceTableAdapters.tbltrans_purchasereturn_itemTableAdapter
        Dim TBLPOC As New DseTransaksiPO_SetOrderBalance.tbltrans_poc_itemDataTable
        Dim DTAPOC As New DseTransaksiPO_SetOrderBalanceTableAdapters.tbltrans_poc_itemTableAdapter
        Dim TBLGIT As New DseTransaksiPO_SetOrderBalance.tbltrans_packinglist_itemDataTable
        Dim DTAGIT As New DseTransaksiPO_SetOrderBalanceTableAdapters.tbltrans_packinglist_itemTableAdapter
        Dim TBLGRN As New DseTransaksiPO_SetOrderBalance.tbltrans_grn_itemDataTable
        Dim DTAGRN As New DseTransaksiPO_SetOrderBalanceTableAdapters.tbltrans_grn_itemTableAdapter
        Dim JmlPO, JmlPR, JmlGIT, JmlGRN, JmlPOC, JMLPOB As Decimal
        Dim JmlTotalGIT, JmlTotalGRN, JmlTotalDP As Decimal

        ''@rosadinnn
        ''@MUaT
        Dim JmlPO_Multi, JmlPR_Multi, JmlGIT_Multi, JmlGRN_Multi, JmlPOC_Multi, JMLPOB_Multi As Decimal
        Dim dtaRatio As New DseTransaksiPOTableAdapters.tbltrans_po_item6TableAdapter
        Dim tblRatio As New DseTransaksiPO.tbltrans_po_item6DataTable
        Dim rowRatio As DseTransaksiPO.tbltrans_po_item6Row
        Dim ratioRef As Decimal = 1
        Dim ratioDef As Decimal = 1

        Dim TBLDP As New DseTransaksiAOP.tbltrans_aopDataTable
        Dim DTADP As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter

        Try
            DTAPO.Connection.ConnectionString = Core_Module.GetConstring
            DTAPR.Connection.ConnectionString = DTAPO.Connection.ConnectionString
            DTAPOC.Connection.ConnectionString = DTAPO.Connection.ConnectionString
            DTAGIT.Connection.ConnectionString = DTAPO.Connection.ConnectionString
            DTAGRN.Connection.ConnectionString = DTAPO.Connection.ConnectionString
            dtaRatio.Connection.ConnectionString = DTAPO.Connection.ConnectionString
            DTADP.Connection.ConnectionString = DTAPO.Connection.ConnectionString

            'fill data table
            JmlTotalGIT = 0
            JmlTotalGRN = 0
            JmlTotalDP = 0

            DTAPO.FillBy_PrimaryPO(TBLPO, vPrimaryPO)
            DTAPR.FillBy_PrimaryPO(TBLPR, vPrimaryPO)
            DTAPOC.FillBy_PrimaryPO(TBLPOC, vPrimaryPO)
            DTAGIT.FillBy_PrimaryPO2(TBLGIT, vPrimaryPO)
            DTAGRN.FillBy_PrimaryPO(TBLGRN, vPrimaryPO)

            ''@rosadinnn
            ''@MUaT
            dtaRatio.Fill(tblRatio, vPrimaryPO)

            Dim dictRatio As New Dictionary(Of Long, DataRow)
            For Each r As DataRow In tblRatio.Rows
                dictRatio(CLng(r("Primary Product"))) = r
            Next

            For Each rowPO In TBLPO.Rows
                JmlPO = rowPO.Qty
                JmlPR = Me.GetJmlPRNewVer(rowPO.Primary_Product, TBLPR, False)
                JmlPOC = Me.GetJmlPOCNewVer(rowPO.Primary_Product, TBLPOC, False)
                JmlGRN = Me.GetJmlGRNNewVer(rowPO.Primary_Product, TBLGRN, False)
                JmlGIT = Me.GetJmlGITNewVer(rowPO.Primary_Product, TBLGIT, False)
                JMLPOB = JmlPO + JmlPR - JmlGIT - JmlPOC - JmlGRN

                ''@rosadinnn
                ''@MUaT
                JmlPO_Multi = rowPO.Qty_Multi
                JmlPR_Multi = Me.GetJmlPRNewVer(rowPO.Primary_Product, TBLPR, True)
                JmlPOC_Multi = Me.GetJmlPOCNewVer(rowPO.Primary_Product, TBLPOC, True)
                JmlGRN_Multi = Me.GetJmlGRNNewVer(rowPO.Primary_Product, TBLGRN, True)
                JmlGIT_Multi = Me.GetJmlGITNewVer(rowPO.Primary_Product, TBLGIT, True)

                ''ubah cara perhitungan POB default(dg menerjemahkan POB multi)
                'tblRatio.DefaultView.RowFilter = "[Primary Product] = " & rowPO.Primary_Product.ToString
                'If tblRatio.DefaultView.Count = 0 Then
                '    MsgBox("Logic error, while trying set POB default(primary product: " & rowPO.Primary_Product.ToString & "). Please contact your software provider.", MsgBoxStyle.Information, "Information")
                '    ratioRef = 1
                '    ratioDef = 1
                'Else
                '    rowRatio = tblRatio.DefaultView(0).Row
                '    ratioRef = rowRatio.Ratio_Reference
                '    ratioDef = rowRatio.Ratio_Default
                'End If

                Dim key As Long = CLng(rowPO.Primary_Product)
                If dictRatio.ContainsKey(key) Then
                    rowRatio = dictRatio(key)
                    ratioRef = rowRatio("Ratio Reference")
                    ratioDef = rowRatio("Ratio Default")
                Else
                    MsgBox("Logic error, primary product: " & key)
                    ratioRef = 1
                    ratioDef = 1
                End If

                JMLPOB_Multi = JmlPO_Multi + JmlPR_Multi - JmlGIT_Multi - JmlPOC_Multi - JmlGRN_Multi
                JMLPOB = Math.Round(CDec(JMLPOB_Multi * ratioDef / ratioRef), CInt(rowPO._Decimal), MidpointRounding.AwayFromZero)

                DTAPO.UpdateQuery_setGIT(JmlGIT, JmlGIT_Multi, rowPO.Primary_PO, rowPO.Primary_Product)
                DTAPO.UpdateQuery_SetPOB(JMLPOB, JMLPOB_Multi, rowPO.Primary_PO, rowPO.Primary_Product)

                JmlTotalGIT += JmlGIT + JmlPOC
                JmlTotalGRN += JmlGRN
            Next

            DTADP.FillBy_PrimaryPO_DP(TBLDP, vPrimaryPO)
            JmlTotalDP = TBLDP.Rows.Count

            If JmlTotalGRN > 0 Or JmlTotalGIT > 0 Or JmlTotalDP > 0 Then
                DTAPO.UpdateQuery(JmlTotalGRN + JmlTotalGIT + JmlTotalDP, vPrimaryPO)
            Else
                DTAPO.UpdateQuery(0, vPrimaryPO)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLPO.Dispose()
            DTAPO.Dispose()
            TBLPR.Dispose()
            DTAPR.Dispose()
            TBLPOC.Dispose()
            DTAPOC.Dispose()
            TBLGIT.Dispose()
            DTAGIT.Dispose()
            TBLGRN.Dispose()
            DTAGRN.Dispose()
            dtaRatio.Dispose()
            DTADP.Dispose()
        End Try
    End Sub
End Class
