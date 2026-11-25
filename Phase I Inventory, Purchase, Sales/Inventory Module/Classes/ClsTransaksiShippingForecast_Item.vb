Option Explicit On
Option Strict Off

Public Class ClsTransaksiShippingForecast_Item
    Public Enum EnumTipeItemPOB As Integer
        POB
        GIT
    End Enum

    Public Enum EnumTipeItemSF As Integer
        SOB
        Purchase_Req
    End Enum

    Public Sub Hapus(ByVal vPrimaryTransaksi As Double)
        Dim DTA1 As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecast_item_pobTableAdapter
        Dim DTA2 As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecast_itemdeliveryTableAdapter
        Dim DTA3 As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecast_itemstockTableAdapter

        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = DTA1.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA2.Connection.ConnectionString

            DTA1.Delete(vPrimaryTransaksi)
            DTA2.Delete(vPrimaryTransaksi)
            DTA3.DeleteQuery(vPrimaryTransaksi)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
        End Try
    End Sub

    Public Sub Simpan(ByVal vPrimaryTransaksi As Double, ByVal vDTItem_Stock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable, _
    ByVal vDTItem_POB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable, _
    ByVal vDTItem_Delivery As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable, ByVal vFrm As FrmTrans_MaterialPlanningDashboard)
        Try
            Me.Simpan_ItemPOB_UsingScript(vDTItem_POB, vPrimaryTransaksi)
            Me.Simpan_ItemDelivery_UsingScript(vDTItem_Delivery, vPrimaryTransaksi)
            Me.Simpan_ItemStock_UsingScript(vDTItem_Stock, vPrimaryTransaksi)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Simpan_ItemStock_UsingScript(ByVal vTBL As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable, ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemStock(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemStock(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable, ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim row As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_shipping_forecast_itemstock` (`primary transaksi`,`primary code no`,`qty`,`primary unit`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTransaksi).ToString + "," + CDbl(.primary_code_no).ToString + "," + CDec(.qty).ToString + "," _
                        + CDbl(.primary_unit).ToString + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Function Simpan_ItemDelivery_UsingScript(ByVal vTBL As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable, ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemDelivery(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemDelivery(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable, ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim row As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_shipping_forecast_itemdelivery` (`primary transaksi`,`primary code no`,`primary so`,`primary cust`,`primary unit`," + _
            "`qty sob`,`so number rec`,`so date rec`,`primary cust rec`,`remark`,`so created by`,`primary priority level`,`promise to deliver`," + _
            "`estimate delivery`,`create do`,`urutan priority level`,`qty available`, `identifier`, `qty ready stock`, `no booking`, `booked`, `qty booked`, `tipe`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        If .so_number_rec <> "" Then
                            If .so_number_rec.Contains("\") = True Then
                                .so_number_rec = .so_number_rec.Replace("\", "\\")
                            End If
                            If .so_number_rec.Contains(vbCrLf) = True Then
                                .so_number_rec = .so_number_rec.Replace(vbCrLf, "\r\n")
                            End If
                            If .so_number_rec.Contains("'") = True Then
                                .so_number_rec = .so_number_rec.Replace("'", "''")
                            End If
                        End If
                        If .so_created_by <> "" Then
                            If .so_created_by.Contains("\") = True Then
                                .so_created_by = .so_created_by.Replace("\", "\\")
                            End If
                            If .so_created_by.Contains(vbCrLf) = True Then
                                .so_created_by = .so_created_by.Replace(vbCrLf, "\r\n")
                            End If
                            If .so_created_by.Contains("'") = True Then
                                .so_created_by = .so_created_by.Replace("'", "''")
                            End If
                        End If
                        If .remark <> "" Then
                            If .remark.Contains("\") = True Then
                                .remark = .remark.Replace("\", "\\")
                            End If
                            If .remark.Contains(vbCrLf) = True Then
                                .remark = .remark.Replace(vbCrLf, "\r\n")
                            End If
                            If .remark.Contains("'") = True Then
                                .remark = .remark.Replace("'", "''")
                            End If
                        End If

                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTransaksi).ToString + "," + CDbl(.primary_code_no).ToString + "," + CDbl(.primary_so).ToString + "," _
                        + CDbl(.primary_cust).ToString + "," + CDbl(.primary_unit).ToString + "," + CDec(.qty_sob).ToString + "," + "'" + .so_number_rec + "'" + "," _
                        + "'" + Format(CDate(.so_date_rec), "yyyy-MM-dd") + "'" + "," + CDbl(.primary_cust_rec).ToString + "," + "'" + .remark + "'" + "," _
                        + "'" + .so_created_by + "'" + "," + CDbl(.primary_priority_level).ToString + "," + "'" + Format(CDate(.promise_to_deliver), "yyyy-MM-dd") + "'" + "," _
                        + "'" + Format(CDate(.estimate_delivery), "yyyy-MM-dd") + "'" + "," + CInt(.create_do).ToString + "," + CInt(.urutan_priority_level).ToString + "," _
                        + CDec(.qty_available).ToString + "," + "'" + .identifier + "'" + "," + CDec(.qty_ready_stock).ToString _
                        + "," + "'" + .no_booking + "'" + "," + CInt(.booked).ToString + "," + CDec(.qty_booked).ToString + "," + CInt(.tipe).ToString + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Function Simpan_ItemPOB_UsingScript(ByVal vTBL As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable, ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemPOB(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemPOB(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable, ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim row As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_shipping_forecast_item_pob` (`primary transaksi`,`primary po`,`primary code no`,`primary supplier`," + _
            "`qty`,`promise to deliver`,`remarks`,`po number rec`,`po date rec`,`primary supplier rec`,`tipe`,`primary packing list`,`packing list no rec`," + _
            "`packing list date`,`delivery delay`, `identifier`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        If .po_number_rec <> "" Then
                            If .po_number_rec.Contains("\") = True Then
                                .po_number_rec = .po_number_rec.Replace("\", "\\")
                            End If
                            If .po_number_rec.Contains(vbCrLf) = True Then
                                .po_number_rec = .po_number_rec.Replace(vbCrLf, "\r\n")
                            End If
                            If .po_number_rec.Contains("'") = True Then
                                .po_number_rec = .po_number_rec.Replace("'", "''")
                            End If
                        End If
                        If .packing_list_no_rec <> "" Then
                            If .packing_list_no_rec.Contains("\") = True Then
                                .packing_list_no_rec = .packing_list_no_rec.Replace("\", "\\")
                            End If
                            If .packing_list_no_rec.Contains(vbCrLf) = True Then
                                .packing_list_no_rec = .packing_list_no_rec.Replace(vbCrLf, "\r\n")
                            End If
                            If .packing_list_no_rec.Contains("'") = True Then
                                .packing_list_no_rec = .packing_list_no_rec.Replace("'", "''")
                            End If
                        End If

                        If .remarks <> "" Then
                            If .remarks.Contains("\") = True Then
                                .remarks = .remarks.Replace("\", "\\")
                            End If
                            If .remarks.Contains(vbCrLf) = True Then
                                .remarks = .remarks.Replace(vbCrLf, "\r\n")
                            End If
                            If .remarks.Contains("'") = True Then
                                .remarks = .remarks.Replace("'", "''")
                            End If
                        End If

                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTransaksi).ToString + "," + CDbl(.primary_po).ToString + "," + CDbl(.primary_code_no).ToString + "," _
                        + CDbl(.primary_supplier).ToString + "," + CDec(.qty).ToString + "," _
                        + "'" + Format(CDate(.promise_to_deliver), "yyyy-MM-dd") + "'" + "," + "'" + .remarks + "'" + "," _
                        + "'" + .po_number_rec + "'" + "," + "'" + Format(CDate(.po_date_rec), "yyyy-MM-dd") + "'" + "," + CDbl(.primary_supplier_rec).ToString + "," _
                        + CInt(.tipe).ToString + "," + CDbl(.primary_packing_list).ToString + "," + "'" + .packing_list_no_rec + "'" + "," + "'" + Format(CDate(.packing_list_date), "yyyy-MM-dd") + "'" + "," _
                        + CInt(.delivery_delay).ToString + "," + "'" + .identifier + "'" + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Public Sub IsiDGV(ByVal vPrimaryTransaksi As Double, ByVal vDTItem_Stock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable, _
    ByVal vDTItem_POB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable, _
    ByVal vDTItem_Delivery As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable)
        Dim DTA_ItemStock As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecast_itemstockTableAdapter
        Dim DTA_ItemDelivery As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecast_itemdeliveryTableAdapter
        Dim DTA_ItemPOB As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecast_item_pobTableAdapter
        Dim LoadRefParty As Boolean = False
        Dim PrimaryMax As Double = 0
        Dim PrimaryMin As Double = Double.MaxValue
        Dim DTAPurcReq As New DseTransaksiShippingForecastTableAdapters.tbltrans_purch_reqTableAdapter
        Dim TBLPurchReq As New DseTransaksiShippingForecast.tbltrans_purch_reqDataTable
        Dim rowPurchReq As DseTransaksiShippingForecast.tbltrans_purch_reqRow
        Dim ObjRefParty As New Core_Module.ClsMasterRefParty

        Try
            DTA_ItemDelivery.Connection.ConnectionString = Core_Module.GetConstring
            DTA_ItemPOB.Connection.ConnectionString = DTA_ItemDelivery.Connection.ConnectionString
            DTA_ItemStock.Connection.ConnectionString = DTA_ItemDelivery.Connection.ConnectionString
            DTAPurcReq.Connection.ConnectionString = DTA_ItemDelivery.Connection.ConnectionString

            DTA_ItemDelivery.FillBy_Primary(vDTItem_Delivery, vPrimaryTransaksi)
            DTA_ItemPOB.FillBy_Primary(vDTItem_POB, vPrimaryTransaksi)
            DTA_ItemStock.FillBy_Primary(vDTItem_Stock, vPrimaryTransaksi)

            For Each rowItemSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow In vDTItem_Delivery.Rows
                If rowItemSOB.tipe = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req Then
                    LoadRefParty = True
                    Exit For
                End If
            Next

            'khusus item delivery, ada perlakukan lanjutan khusus terkait dengan primary so yang sudah di hapuskan
            For Each rowItemSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow In vDTItem_Delivery.Rows
                If rowItemSOB.tipe = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB Then
                    If rowItemSOB.IsCust_PO_NoNull = True Or rowItemSOB.Isso_numberNull = True Or rowItemSOB.Isso_dateNull = True Then
                        rowItemSOB.Cust_PO_No = ""
                        rowItemSOB.so_number = rowItemSOB.so_number_rec
                        rowItemSOB.so_date = rowItemSOB.so_date_rec
                    End If
                End If
            Next
            If LoadRefParty = True Then
                For Each rowItemSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow In vDTItem_Delivery.Rows
                    If rowItemSOB.tipe = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req Then
                        If PrimaryMax < rowItemSOB.primary_so Then
                            PrimaryMax = rowItemSOB.primary_so
                        End If
                        If PrimaryMin > rowItemSOB.primary_so Then
                            PrimaryMin = rowItemSOB.primary_so
                        End If
                    End If
                Next

                DTAPurcReq.Fill(TBLPurchReq, PrimaryMax, PrimaryMin)
                For Each rowItemSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow In vDTItem_Delivery.Rows
                    If rowItemSOB.tipe = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req Then
                        TBLPurchReq.DefaultView.RowFilter = "primary = " + rowItemSOB.primary_so.ToString
                        If TBLPurchReq.DefaultView.Count > 0 Then
                            rowPurchReq = TBLPurchReq.DefaultView(0).Row
                            'purchase req
                            rowItemSOB.Cust_PO_No = ""
                            rowItemSOB.so_number = rowPurchReq.no_transaksi
                            rowItemSOB.so_date = rowPurchReq.tanggal
                            rowItemSOB.customer_name = rowPurchReq.Ref_Party
                        Else
                            'ambil dari history
                            rowItemSOB.Cust_PO_No = ""
                            rowItemSOB.so_number = rowItemSOB.so_number_rec
                            rowItemSOB.so_date = rowItemSOB.so_date_rec
                            rowItemSOB.customer_name = ObjRefParty.GetDesc(rowItemSOB.primary_cust_rec)
                        End If

                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPurcReq.Dispose()
            TBLPurchReq.Dispose()
            DTA_ItemStock.Dispose()
            DTA_ItemDelivery.Dispose()
            DTA_ItemPOB.Dispose()
            ObjRefParty = Nothing
        End Try
    End Sub
End Class
