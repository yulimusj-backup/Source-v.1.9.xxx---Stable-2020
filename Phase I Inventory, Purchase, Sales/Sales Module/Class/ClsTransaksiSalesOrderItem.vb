Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesOrderItem
    Public Sub Set_Order_Balance(ByVal vListPrimarySO As List(Of Double))
        Dim PrimarySO As Double
        Try
            For Each PrimarySO In vListPrimarySO
                Me.Set_Order_Balance(PrimarySO)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Set_Order_Balance(ByVal vPrimarySO As Double)
        Dim TBLSO As New DseTransaksiSO_OrderBalance.tbltrans_so_itemDataTable
        Dim TBLDO As New DseTransaksiSO_OrderBalance.tbltrans_do_itemDataTable
        Dim TBLSOC As New DseTransaksiSO_OrderBalance.DataTable1DataTable
        Dim TBLSR As New DseTransaksiSO_OrderBalance.tbltrans_salesreturn_itemDataTable
        Dim TBLDOLoss As New DseTransaksiSO_OrderBalance.tbltrans_do_loss_itemDataTable
        Dim DTASO As New DseTransaksiSO_OrderBalanceTableAdapters.tbltrans_so_itemTableAdapter
        Dim DTADO As New DseTransaksiSO_OrderBalanceTableAdapters.tbltrans_do_itemTableAdapter
        Dim DTASOC As New DseTransaksiSO_OrderBalanceTableAdapters.DataTable1TableAdapter
        Dim DTASR As New DseTransaksiSO_OrderBalanceTableAdapters.tbltrans_salesreturn_itemTableAdapter
        Dim DTADOLoss As New DseTransaksiSO_OrderBalanceTableAdapters.tbltrans_do_loss_itemTableAdapter
        Dim rowSO As DseTransaksiSO_OrderBalance.tbltrans_so_itemRow
        Dim SOB As Decimal
        Dim JmlDO_Multi, JmlSOC_Multi, JmlSR_Multi, JmlDOLoss_Multi, SOB_Multi As Decimal

        Try
            'fill database
            DTASO.FillBy_PrimarySO(TBLSO, vPrimarySO)
            DTADO.FillBy_PrimarySO(TBLDO, vPrimarySO)
            DTASOC.FillBy_PrimarySO(TBLSOC, vPrimarySO)
            DTASR.FillBy(TBLSR, vPrimarySO)
            DTADOLoss.FillBy_PrimarySO(TBLDOLoss, vPrimarySO)

            For Each rowSO In TBLSO.Rows
                With rowSO
                    JmlDOLoss_Multi = Me.GetJmlDOLoss_Multi(.Primary_Transaksi, .Primary_Product, TBLDOLoss)
                    JmlDO_Multi = Me.GetJmlDO_Multi(.Primary_Transaksi, .Primary_Product, TBLDO)
                    JmlSR_Multi = Me.GetJmlSR_Multi(.Primary_Product, .Primary_Transaksi, TBLSR)
                    JmlSOC_Multi = Me.GetJmlSOC_Multi(.Primary_Transaksi, .Primary_Product, TBLSOC)
                    SOB_Multi = .Qty_Multi - JmlDO_Multi - JmlSOC_Multi + JmlSR_Multi - JmlDOLoss_Multi
                    If .Ratio_Default = 0 Then
                        .Ratio_Default = 1
                    End If
                    If .Ratio_Reference = 0 Then
                        .Ratio_Reference = 1
                    End If
                    SOB = Math.Round(CDec(SOB_Multi * .Ratio_Default / .Ratio_Reference), CInt(._Decimal), MidpointRounding.AwayFromZero)
                    DTASO.UpdateQuery_SetSOB(SOB, SOB_Multi, .Primary)
                End With
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLSO.Dispose()
            TBLDO.Dispose()
            TBLSOC.Dispose()
            TBLSR.Dispose()
            TBLDOLoss.Dispose()
            DTASO.Dispose()
            DTADO.Dispose()
            DTASOC.Dispose()
            DTASR.Dispose()
            DTADOLoss.Dispose()
        End Try
    End Sub

    Private Function GetJmlDOLoss_Multi(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, _
    ByVal TBLDOLoss As DseTransaksiSO_OrderBalance.tbltrans_do_loss_itemDataTable) As Decimal
        Dim rowView As DataRowView
        Dim rowDOLoss As DseTransaksiSO_OrderBalance.tbltrans_do_loss_itemRow
        Dim Filter As String = "[Primary SO] =" + vPrimarySO.ToString _
        + " AND [Primary Product] =" + vPrimaryProduct.ToString
        Dim Dummy As Decimal = 0

        Try
            TBLDOLoss.DefaultView.RowFilter = Filter
            If TBLDOLoss.DefaultView.Count > 0 Then
                For Each rowView In TBLDOLoss.DefaultView
                    rowDOLoss = rowView.Row
                    Dummy += rowDOLoss.Total_Qty_Multi
                Next

                Return Dummy
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDO_Multi(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, _
    ByVal TBLDO As DseTransaksiSO_OrderBalance.tbltrans_do_itemDataTable) As Decimal
        Dim rowView As DataRowView
        Dim rowDO As DseTransaksiSO_OrderBalance.tbltrans_do_itemRow
        Dim Filter As String = "[Primary SO] =" + vPrimarySO.ToString _
        + " AND [Primary Product] =" + vPrimaryProduct.ToString
        Dim Dummy As Decimal = 0

        Try
            TBLDO.DefaultView.RowFilter = Filter
            If TBLDO.DefaultView.Count > 0 Then
                For Each rowView In TBLDO.DefaultView
                    rowDO = rowView.Row
                    Dummy += rowDO.Total_Qty_Multi
                Next

                Return Dummy
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlSOC_Multi(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, _
    ByVal TBLSOC As DseTransaksiSO_OrderBalance.DataTable1DataTable) As Decimal
        Dim rowView As DataRowView
        Dim rowSOC As DseTransaksiSO_OrderBalance.DataTable1Row
        Dim Filter As String = "[Primary SO] =" + vPrimarySO.ToString _
        + " AND [Primary Product] =" + vPrimaryProduct.ToString
        Dim Dummy As Decimal = 0

        Try
            TBLSOC.DefaultView.RowFilter = Filter
            If TBLSOC.DefaultView.Count > 0 Then
                For Each rowView In TBLSOC.DefaultView
                    rowSOC = rowView.Row
                    Dummy += rowSOC.Total_Qty_Multi
                Next

                Return Dummy
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlSR_Multi(ByVal vPrimaryProduct As Double, ByVal vPrimarySO As Double, _
    ByVal TBLSR As DseTransaksiSO_OrderBalance.tbltrans_salesreturn_itemDataTable) As Decimal
        Dim rowView As DataRowView
        Dim rowSR As DseTransaksiSO_OrderBalance.tbltrans_salesreturn_itemRow
        Dim Filter As String = "[Primary SO] =" + vPrimarySO.ToString _
        + " AND [Primary Product] =" + vPrimaryProduct.ToString
        Dim Dummy As Decimal = 0

        Try
            TBLSR.DefaultView.RowFilter = Filter
            If TBLSR.DefaultView.Count > 0 Then
                For Each rowView In TBLSR.DefaultView
                    rowSR = rowView.Row
                    Dummy += rowSR.Total_Qty_Multi
                Next

                Return Dummy
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetRemarks(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double) As String
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetRemarks(vPrimarySO, vPrimaryProduct)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetRemarks(vPrimarySO, vPrimaryProduct)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    

    Public Sub IsikanTBLProduct(ByVal vPrimarySalesOrder As Double, _
    ByVal TBL As DseTransaksiSO.laporan_sales_orderbalance_view9aDataTable)
        Dim DTA As New DseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9aTableAdapter

        Try
            DTA.FillBy_PrimarySO(TBL, vPrimarySalesOrder)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLSO(ByVal vPrimaryCustomer As Double, ByVal TBL As DseTransaksiSO.DataTable3DataTable, _
   ByVal cbo As ComboBox, ByVal vPrimaryProduct As Double)
        Dim DTA As New DseTransaksiSOTableAdapters.DataTable3TableAdapter

        Try
            DTA.FillBy_PrimarySO_PrimaryProduct(TBL, vPrimaryProduct, vPrimaryCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLSO2(ByVal vPrimaryCustomer As Double, ByVal TBL As DseTransaksiSO.laporan_sales_orderbalance_view9cDataTable, _
       ByVal cbo As ComboBox, ByVal vPrimaryProduct As Double)
        Dim DTA As New DseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9cTableAdapter

        Try
            DTA.FillBy_PrimaryCustomer_PrimaryProduct1(TBL, vPrimaryCustomer, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLSO(ByVal vPrimaryCustomer As Double, ByVal TBL As DseTransaksiSO.DataTable3DataTable, _
    ByVal cbo As ComboBox)
        Dim DTA As New DseTransaksiSOTableAdapters.DataTable3TableAdapter

        Try
            DTA.FillBy_PrimaryCustomer(TBL, vPrimaryCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLSO3_Tolerance(ByVal vPrimaryCustomer As Double, ByVal TBL As DseTransaksiSO.laporan_sales_orderbalance_view9dDataTable, _
        ByVal cbo As ComboBox)
        Dim DTA As New DseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9dTableAdapter

        Try
            DTA.FillBy_PrimaryCust_Tolerance(TBL, vPrimaryCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLSO3_Default(ByVal vPrimaryCustomer As Double, ByVal TBL As DseTransaksiSO.laporan_sales_orderbalance_view9dDataTable, _
        ByVal cbo As ComboBox)
        Dim DTA As New DseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9dTableAdapter

        Try
            DTA.FillBy_PrimaryCust(TBL, vPrimaryCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLProductByCustomer(ByVal vPrimaryCustomer As Double, ByVal TBL As DseTransaksiSO.laporan_sales_orderbalance_view9cDataTable, _
    ByVal cbo As ComboBox)
        Dim DTA As New DseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9cTableAdapter

        Try
            DTA.FillBy_PrimaryCustomer(TBL, vPrimaryCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLProductByCustomer2(ByVal vPrimaryCustomer As Double, ByVal TBL As DseTransaksiSO.laporan_sales_orderbalance_view9cDataTable, _
    ByVal cbo As ComboBox)
        Dim DTA As New DseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9cTableAdapter

        Try
            DTA.FillBy_PrimaryCustomer(TBL, vPrimaryCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLProduct2_Default(ByVal vPrimarySalesOrder As Double, ByVal TBL As DseTransaksiSO.laporan_sales_orderbalance_view9dDataTable)
        Dim DTA As New DseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9dTableAdapter

        Try
            DTA.FillBy_PrimarySO(TBL, vPrimarySalesOrder)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLProduct2_Tolerance(ByVal vPrimarySalesOrder As Double, ByVal TBL As DseTransaksiSO.laporan_sales_orderbalance_view9dDataTable)
        Dim DTA As New DseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9dTableAdapter

        Try
            DTA.FillBy_PrimarySO_Tolerance(TBL, vPrimarySalesOrder)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLProduct(ByVal vPrimarySalesOrder As Double, ByVal TBL As DseTransaksiSO.DataTable3DataTable, _
    ByVal cbo As ComboBox)
        Dim DTA As New DseTransaksiSOTableAdapters.DataTable3TableAdapter

        Try
            DTA.FillBy_PrimarySO(TBL, vPrimarySalesOrder)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function IsiDGV(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal vDT_ItemDelivery As DseTransaksiSO.tbltrans_so_itemdeliveryDataTable, ByVal vDefPrecCurSO As Integer) As Boolean
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_so_item1TableAdapter
        Dim TBL As New DseTransaksiSO.tbltrans_so_item1DataTable
        Dim DTAItemDelivery As New DseTransaksiSOTableAdapters.tbltrans_so_itemdeliveryTableAdapter
        Dim FrmSO As FrmTransaksiSO = vDGVItem.FindForm

        Dim row As DseTransaksiSO.tbltrans_so_item1Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PriceCondition As String = ""
        Dim DeliveryDate_Str As String = ""
        Dim rowDGV As DataGridViewRow

        Dim DefPrecMulti As Integer = 0
        Dim ObjProduct As New Core_Module.clsMasterProduct

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemDelivery.Connection.ConnectionString = DTA.Connection.ConnectionString

            'kosongkan dgv
            vDGVItem.Rows.Clear()
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        If .Price_Exclude = True Then
                            PriceCondition = "Exclude"
                        Else
                            PriceCondition = "Include"
                        End If
                        If CBool(.Delivery_Date_Cont) = True Then
                            If .Delivery_Date = #1/1/1950# Then
                                DeliveryDate_Str = "Multiple"
                            Else
                                DeliveryDate_Str = Format(.Delivery_Date, "dd MMM yyyy")
                            End If
                        Else
                            DeliveryDate_Str = "-"
                        End If
                        vDGVItem.Rows.Add()
                        rowDGV = vDGVItem.Rows(vDGVItem.RowCount - 1)
                        rowDGV.Cells(FrmSO.ColAmount.Index).Value = ObjCD.DigitGrouping(.Amount, vDefPrecCurSO)
                        rowDGV.Cells(FrmSO.ColCodeNo.Index).Value = .Product_Code
                        rowDGV.Cells(FrmSO.ColDeliveryCont.Index).Value = CBool(.Delivery_Date_Cont)
                        rowDGV.Cells(FrmSO.ColDeliveryDate.Index).Value = DeliveryDate_Str
                        rowDGV.Cells(FrmSO.ColDeliveryHistory.Index).Value = ""
                        rowDGV.Cells(FrmSO.ColDescription.Index).Value = .Description
                        rowDGV.Cells(FrmSO.ColDiscAmount.Index).Value = ObjCD.DigitGrouping(.Disc_Amount, vDefPrecCurSO)
                        rowDGV.Cells(FrmSO.ColDiscount.Index).Value = ObjCD.DigitGrouping(.Discount, 2)
                        rowDGV.Cells(FrmSO.ColIdentifier.Index).Value = .identifier
                        rowDGV.Cells(FrmSO.ColNettPrice.Index).Value = ObjCD.DigitGrouping(.Nett_Price, vDefPrecCurSO)
                        rowDGV.Cells(FrmSO.ColNo.Index).Value = vDGVItem.Rows.Count
                        rowDGV.Cells(FrmSO.ColPriceCond.Index).Value = PriceCondition
                        rowDGV.Cells(FrmSO.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rowDGV.Cells(FrmSO.ColPrimaryUnits.Index).Value = .Primary_Units
                        rowDGV.Cells(FrmSO.ColQty.Index).Value = ObjCD.DigitGrouping(.Qty, ._Decimal)
                        rowDGV.Cells(FrmSO.ColRemark.Index).Value = .Remarks
                        rowDGV.Cells(FrmSO.ColSuppDesc.Index).Value = ""
                        rowDGV.Cells(FrmSO.ColSupplierCodeNo.Index).Value = ""
                        rowDGV.Cells(FrmSO.ColUnit.Index).Value = .Unit_Code
                        rowDGV.Cells(FrmSO.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(.Unit_Price, vDefPrecCurSO)
                        rowDGV.Cells(FrmSO.ColUseSecondary.Index).Value = CBool(.Use_Secondary)
                        rowDGV.Cells(FrmSO.ColDiscFromPerct.Index).Value = .Disc_From_Perct

                        DefPrecMulti = ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)
                        rowDGV.Cells(FrmSO.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(.Qty_Multi, DefPrecMulti)
                        rowDGV.Cells(FrmSO.ColUnitsMulti.Index).Value = .Units_Code_Multi
                        rowDGV.Cells(FrmSO.ColPrimaryUnitsMulti.Index).Value = .Primary_Units_Multi
                        rowDGV.Cells(FrmSO.ColToleranceDown.Index).Value = ObjCD.DigitGrouping(.Tolerance_Down, DefPrecMulti)
                        rowDGV.Cells(FrmSO.ColToleranceDownPerc.Index).Value = ObjCD.DigitGrouping(.Tolerance_Down_Perc, 2)
                        rowDGV.Cells(FrmSO.ColToleranceDownFromPerc.Index).Value = ObjCD.DigitGrouping(.Tolerance_Down_From_Perc, 0)
                        rowDGV.Cells(FrmSO.ColToleranceUp.Index).Value = ObjCD.DigitGrouping(.Tolerance_Up, DefPrecMulti)
                        rowDGV.Cells(FrmSO.ColToleranceUpPerc.Index).Value = ObjCD.DigitGrouping(.Tolerance_Up_Perc, 2)
                        rowDGV.Cells(FrmSO.ColToleranceUpFromPerc.Index).Value = ObjCD.DigitGrouping(.Tolerance_Up_From_Perc, 0)
                        rowDGV.Cells(FrmSO.ColRatioReference.Index).Value = .Ratio_Reference
                        rowDGV.Cells(FrmSO.ColRatioDefault.Index).Value = .Ratio_Default
                        rowDGV.Cells(FrmSO.ColPriceByMultiUnit.Index).Value = .price_by_multi_unit
                        If .price_by_multi_unit = 1 Then
                            rowDGV.Cells(FrmSO.ColPriceByMultiUnit_Str.Index).Value = "/" + .Units_Code_Multi
                        Else
                            rowDGV.Cells(FrmSO.ColPriceByMultiUnit_Str.Index).Value = "/" + .Unit_Code
                        End If
                        'rowDGV.Cells(FrmSO.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti) & " " & .Units_Code_Multi & "=" & ObjCD.DigitGrouping(.Ratio_Default, ._Decimal) & " " & .Unit_Code

                        'perbaikan hanya pada kolom dgv ratio desc
                        If .Primary_Units <> .Primary_Units_Multi Then
                            rowDGV.Cells(FrmSO.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti) & " " & .Units_Code_Multi & "=" & ObjCD.DigitGrouping(.Ratio_Default, ._Decimal) & " " & .Unit_Code
                        Else
                            rowDGV.Cells(FrmSO.ColRatioDesc.Index).Value = "-"
                        End If
                    End With
                Next
            End If

            'set nomer
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            vDT_ItemDelivery.Rows.Clear()
            DTAItemDelivery.FillBy_Primary(vDT_ItemDelivery, vPrimaryTransaksi)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetJml_Default(ByVal vPrimaryProduct As Double, ByVal vPrimarySO As Double) As Decimal
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetJml(vPrimaryProduct, vPrimarySO).HasValue = True Then
                Return DTA.ScalarQuery_GetJml(vPrimaryProduct, vPrimarySO).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_Multi(ByVal vPrimaryProduct As Double, ByVal vPrimarySO As Double) As Decimal
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetJml_Multi(vPrimaryProduct, vPrimarySO).HasValue = True Then
                Return DTA.ScalarQuery_GetJml_Multi(vPrimaryProduct, vPrimarySO).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiDGVSalesOrderBalance(ByVal vPrimaryProduct As Double, ByVal vTBL As DseTransaksiSO_OrderBalance.tbltrans_so_item1DataTable, _
    ByVal vDGV As DataGridView)
        Dim DTA As New DseTransaksiSO_OrderBalanceTableAdapters.tbltrans_so_item1TableAdapter
        Dim row As DseTransaksiSO_OrderBalance.tbltrans_so_item1Row
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGV.DataSource = Nothing

            If vPrimaryProduct = -1 Then
                DTA.Fill(vTBL)
            Else
                DTA.FillBy_PrimaryProduct(vTBL, vPrimaryProduct)
            End If

            For Each row In vTBL.Rows
                If row.Delivery_Date = #1/1/1950# Then
                    row.Promise_To_Deliver = "-"
                Else
                    row.Promise_To_Deliver = Format(row.Delivery_Date, "dd MMM yyyy")
                End If
                row.Order_Balance = ObjCD.DigitGrouping(row.SOB, row._Decimal)
            Next

            vDGV.DataSource = vTBL.DefaultView
            vDGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
            vDGV.Columns(7).Visible = False
            vDGV.Columns(8).Visible = False
            vDGV.Columns(9).Visible = False
            vDGV.Columns(10).Visible = False
            vDGV.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            vDGV.Columns(11).SortMode = DataGridViewColumnSortMode.Programmatic
            vDGV.Columns(12).SortMode = DataGridViewColumnSortMode.Programmatic
            If vPrimaryProduct <> -1 Then
                vDGV.Columns(5).Visible = False
                vDGV.Columns(6).Visible = False
            End If

            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Sub

    Public Function GetSalesOrderBalance(ByVal vPrimaryProduct As Double) _
    As Decimal
        Dim DTA As New DseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9aTableAdapter

        Try
            If DTA.ScalarQuery_GetOB(vPrimaryProduct).HasValue = True Then
                Return DTA.ScalarQuery_GetOB(vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSalesOrderBalance_Default(ByVal vprimarySO As Double, ByVal vPrimaryProduct As Double) As Decimal
        Dim DTA As New DseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9aTableAdapter
        Try
            If DTA.ScalarQuery_GetOrderBalance(vprimarySO, vPrimaryProduct).HasValue = True Then
                'MsgBox(DTA.ScalarQuery_GetOrderBalance(vprimarySO, vPrimaryProduct).Value.ToString)
                Return DTA.ScalarQuery_GetOrderBalance(vprimarySO, vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSalesOrderBalance_Multi(ByVal vprimarySO As Double, ByVal vPrimaryProduct As Double) As Decimal
        Dim DTA As New DseTransaksiSOTableAdapters.laporan_sales_orderbalance_view9aTableAdapter
        Try
            If DTA.ScalarQuery_GetOrderBalance_Multi(vprimarySO, vPrimaryProduct).HasValue = True Then
                'MsgBox(DTA.ScalarQuery_GetOrderBalance(vprimarySO, vPrimaryProduct).Value.ToString)
                Return DTA.ScalarQuery_GetOrderBalance_Multi(vprimarySO, vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItem As DataGridView, _
    ByVal vPrimaryTransaksi As Double, ByVal vFrm As FrmTransaksiSO, ByVal vDT_ItemDelivery As DseTransaksiSO.tbltrans_so_itemdeliveryDataTable) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter
        Dim DelDateCont, UseSecondary As Byte
        Dim DelDate As Date
        Dim PriceCondition As Integer = 0
        Dim DTAItemDelivery As New DseTransaksiSOTableAdapters.tbltrans_so_itemdeliveryTableAdapter
        Dim rowItemDelivery As DseTransaksiSO.tbltrans_so_itemdeliveryRow
        Dim listPrimaryProduct As New List(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemDelivery.Connection.ConnectionString = DTA.Connection.ConnectionString

            For Each row In vdgvItem.Rows
                'simpan ke basis data item
                If row.Cells(vFrm.ColPriceCond.Index).Value = "Exclude" Then
                    PriceCondition = 1
                Else
                    PriceCondition = 0
                End If
                If row.Cells(vFrm.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                    DelDateCont = 1
                    If IsDate(row.Cells(vFrm.ColDeliveryDate.Index).Value) = True Then
                        DelDate = row.Cells(vFrm.ColDeliveryDate.Index).Value
                    Else
                        DelDate = #1/1/1950#
                    End If
                Else
                    DelDateCont = 0
                    DelDate = #1/1/1950#
                End If
                If row.Cells(vFrm.ColUseSecondary.Index).Value = Boolean.TrueString Then
                    UseSecondary = 1
                Else
                    UseSecondary = 0
                End If

                DTA.InsertQuery(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryCodeNo.Index).Value, row.Cells(vFrm.ColQty.Index).Value, row.Cells(vFrm.ColPrimaryUnits.Index).Value, _
                DelDateCont, DelDate, row.Cells(vFrm.ColUnitPrice.Index).Value, row.Cells(vFrm.ColDiscount.Index).Value, _
                row.Cells(vFrm.ColNettPrice.Index).Value, row.Cells(vFrm.ColAmount.Index).Value, row.Cells(vFrm.ColRemark.Index).Value, _
                PriceCondition, UseSecondary, row.Cells(vFrm.ColIdentifier.Index).Value, row.Cells(vFrm.ColDiscAmount.Index).Value, _
                row.Cells(vFrm.ColDiscFromPerct.Index).Value, _
                row.Cells(vFrm.ColQtyMulti.Index).Value, row.Cells(vFrm.ColPrimaryUnitsMulti.Index).Value, row.Cells(vFrm.ColRatioReference.Index).Value, _
                row.Cells(vFrm.ColRatioDefault.Index).Value, row.Cells(vFrm.ColToleranceDown.Index).Value, row.Cells(vFrm.ColToleranceDownPerc.Index).Value, _
                row.Cells(vFrm.ColToleranceDownFromPerc.Index).Value, row.Cells(vFrm.ColToleranceUp.Index).Value, row.Cells(vFrm.ColToleranceUpPerc.Index).Value, _
                row.Cells(vFrm.ColToleranceDownFromPerc.Index).Value, row.Cells(vFrm.ColQtySpare.Index).Value, row.Cells(vFrm.ColPriceByMultiUnit.Index).Value)

                'DTA.UpdateQuery_SetTblMasterProduct(row.Cells(vFrm.ColPriceByMultiUnit.Index).Value, row.Cells(vFrm.ColPrimaryCodeNo.Index).Value)

                If listPrimaryProduct.Contains(row.Cells(vFrm.ColPrimaryCodeNo.Index).Value) = False Then
                    listPrimaryProduct.Add(row.Cells(vFrm.ColPrimaryCodeNo.Index).Value)
                    DTA.UpdateQuery_SetTblMasterProduct(row.Cells(vFrm.ColPriceByMultiUnit.Index).Value, row.Cells(vFrm.ColPrimaryCodeNo.Index).Value)
                End If
            Next

            For Each rowItemDelivery In vDT_ItemDelivery.Rows
                With rowItemDelivery
                    If .identifier.StartsWith("-1") Then
                        .identifier = vPrimaryTransaksi.ToString + .identifier.Substring(2)
                    End If
                    DTAItemDelivery.InsertQuery(vPrimaryTransaksi, .identifier, .delivery_qty, .delivery_date, .remarks, .delivery_priority, .delivery_qty_multi)
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
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter
        Dim DTA_ItemDelivery As New DseTransaksiSOTableAdapters.tbltrans_so_itemdeliveryTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            DTA_ItemDelivery.Dispose()
        End Try
    End Function
End Class
