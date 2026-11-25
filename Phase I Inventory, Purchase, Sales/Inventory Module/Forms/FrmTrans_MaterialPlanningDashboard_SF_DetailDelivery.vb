Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_SF_DetailDelivery
    Public vTBLItemStock As New DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable
    Public vTBLItemDelivery As New DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable
    Public vTBLItemPOB As New DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable
    Public vPrimaryTarget As Double = -1
    Public vPrimaryCodeNoTarget As Double = -1
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public DataChanged As Boolean = False
    Public vReadOnly As Boolean = False
    Public UserCode As String = ""
    Public vPrimaryStockInPurchaseReference As Double = -1
    Public vFromStockIn As Boolean = False
    Public TBL_ProductQty1 As New DseTransaksiShippingForecast.tblmaster_productqtyDataTable

    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjProd As New Core_Module.clsMasterProduct
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjPriorityLevel As New Core_Module.ClsMasterPriorityLevel
    Dim ObjSupp As New Core_Module.clsMasterSupplier
    Dim Buffer As String = ""
    Dim ObjPOItem As New ClsTransaksiPOItem
    Dim ObjSOItem As New ClsTransaksiSalesOrder
    Dim ObjPurcReq As New ClsTransaksiPR
    Dim ObjPO As New ClsTransaksiPO
    Dim ObjPackingList As New ClsTransaksiPackingList
    Dim TipeProduct As Core_Module.clsMasterProduct.EnumTipeProduct

    Private Enum EnumTipeItem As Integer
        Stock
        POB
        SOB
    End Enum

    Private Sub TambahkanItemDelivery(ByVal vPrimary As Double, ByVal DefPrec As Integer)
        Dim rowDGV As DataGridViewRow
        Dim rowDel As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim QtyBefore As Decimal = 0
        Dim NoSO As String = ""

        Try
            For Each rowDel In Me.vTBLItemDelivery.Rows
                If rowDel.primary = vPrimary Then
                    'tambahkan baris ke dalam dgv, untuk begining balance item ready stock
                    Me.dgvData.Rows.Add()
                    rowDGV = Me.dgvData.Rows(Me.dgvData.Rows.Count - 1)
                    If rowDGV.Index = 0 Then
                        QtyBefore = 0
                    Else
                        If IsNumeric(Me.dgvData.Item(Me.ColBalance.Index, rowDGV.Index - 1).Value) = False Then
                            QtyBefore = 0
                        Else
                            QtyBefore = CDec(Me.dgvData.Item(Me.ColBalance.Index, rowDGV.Index - 1).Value)
                        End If
                    End If

                    If QtyBefore - rowDel.qty_available <> 0 Then
                        rowDGV.Cells(Me.ColBalance.Index).Value = ObjCD.DigitGrouping(QtyBefore - rowDel.qty_available, DefPrec)
                    Else
                        rowDGV.Cells(Me.ColBalance.Index).Value = "-"
                    End If
                    If rowDel.estimate_delivery = #1/1/1950# Then
                        rowDGV.Cells(Me.ColDate.Index).Value = "Today"
                    ElseIf rowDel.estimate_delivery = #1/1/2200# Then
                        rowDGV.Cells(Me.ColDate.Index).Value = "n/a"
                    Else
                        rowDGV.Cells(Me.ColDate.Index).Value = Format(rowDel.estimate_delivery, "dd MMM yyyy")
                    End If
                    If rowDel.tipe = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB Then
                        rowDGV.Cells(Me.ColDesc.Index).Value = "Sales to " + rowDel.customer_name
                        NoSO = Me.ObjSOItem.GetNoSO(rowDel.primary_so)
                    Else
                        rowDGV.Cells(Me.ColDesc.Index).Value = "Purc. Request by " + rowDel.customer_name
                        NoSO = Me.ObjPurcReq.GetNoTransaksi(rowDel.primary_so)
                    End If
                    rowDGV.Cells(Me.ColTipeInt.Index).Value = rowDel.tipe
                    rowDGV.Cells(Me.ColEdit.Index).Value = "Edit"
                    rowDGV.Cells(Me.ColCreatedBy.Index).Value = rowDel.so_created_by
                    rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvData.Rows.Count
                    rowDGV.Cells(Me.ColPrimary.Index).Value = rowDel.primary
                    rowDGV.Cells(Me.ColPrimaryPriorityLevel.Index).Value = rowDel.primary_priority_level
                    rowDGV.Cells(Me.ColPriorityLevel.Index).Value = Me.ObjPriorityLevel.GetDesc(rowDel.primary_priority_level)
                    rowDGV.Cells(Me.ColQtyIn.Index).Value = "-"
                    If rowDel.qty_available > 0 Then
                        rowDGV.Cells(Me.ColQtyOut.Index).Value = ObjCD.DigitGrouping(rowDel.qty_available, DefPrec)
                    Else
                        rowDGV.Cells(Me.ColQtyOut.Index).Value = "-"
                    End If
                    rowDGV.Cells(Me.ColTipeItem.Index).Value = CInt(FrmTrans_MaterialPlanningDashboard_SF_DetailDelivery.EnumTipeItem.SOB)
                    If NoSO <> "" Then
                        rowDGV.Cells(Me.ColDocNo.Index).Value = NoSO
                    Else
                        rowDGV.Cells(Me.ColDocNo.Index).Value = rowDel.so_number
                    End If
                    If rowDel.tipe_product = Core_Module.clsMasterProduct.EnumTipeProduct.Goods Then
                        If rowDel.estimate_delivery > rowDel.promise_to_deliver Or rowDel.qty_available < rowDel.qty_sob Then
                            rowDGV.Cells(Me.ColIsLate.Index).Value = Boolean.TrueString
                            rowDGV.DefaultCellStyle.ForeColor = Color.Red
                        Else
                            rowDGV.Cells(Me.ColIsLate.Index).Value = Boolean.FalseString
                        End If
                    Else
                        rowDGV.Cells(Me.ColIsLate.Index).Value = Boolean.FalseString
                    End If

                    If rowDel.promise_to_deliver = #1/1/2200# Then
                        rowDGV.Cells(Me.ColPromiseToDeliver.Index).Value = "n/a"
                    Else
                        rowDGV.Cells(Me.ColPromiseToDeliver.Index).Value = Format(rowDel.promise_to_deliver, "dd MMM yyyy")
                    End If
                    rowDGV.Cells(Me.ColQtyPromiseToDeliver.Index).Value = ObjCD.DigitGrouping(rowDel.qty_sob, DefPrec)
                    If rowDel.no_booking <> "" Then
                        rowDGV.Cells(Me.ColBookingNo.Index).Value = rowDel.no_booking
                        rowDGV.Cells(Me.ColBookQty.Index).Value = ObjCD.DigitGrouping(rowDel.qty_booked, DefPrec)
                    Else
                        rowDGV.Cells(Me.ColBookingNo.Index).Value = ""
                        rowDGV.Cells(Me.ColBookQty.Index).Value = ""
                    End If

                    If rowDel.tipe_product = Core_Module.clsMasterProduct.EnumTipeProduct.Services Then
                        rowDGV.Cells(Me.ColStatus.Index).Value = "Ready"
                    Else
                        If rowDel.qty_available > 0 Then
                            If rowDel.qty_sob <= rowDel.qty_available Then
                                If rowDel.qty_ready_stock >= rowDel.qty_available Then
                                    rowDGV.Cells(Me.ColStatus.Index).Value = "Ready"
                                ElseIf rowDel.qty_ready_stock > 0 Then
                                    rowDGV.Cells(Me.ColStatus.Index).Value = "Partial"
                                Else
                                    rowDGV.Cells(Me.ColStatus.Index).Value = "Waiting"
                                End If
                            Else
                                If rowDel.qty_ready_stock >= rowDel.qty_available Then
                                    rowDGV.Cells(Me.ColStatus.Index).Value = "Partial"
                                ElseIf rowDel.qty_ready_stock > 0 Then
                                    rowDGV.Cells(Me.ColStatus.Index).Value = "Partial"
                                Else
                                    rowDGV.Cells(Me.ColStatus.Index).Value = "Waiting"
                                End If
                            End If
                        Else
                            rowDGV.Cells(Me.ColStatus.Index).Value = "Waiting"
                        End If
                    End If
                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahSudahAda_Purchase(ByVal vPrimary As Double) As Boolean
        Try
            For Each rowDGV As DataGridViewRow In Me.dgvData.Rows
                If rowDGV.Cells(Me.ColPrimary.Index).Value = vPrimary And _
                rowDGV.Cells(Me.ColTipeItem.Index).Value = CInt(FrmTrans_MaterialPlanningDashboard_SF_DetailDelivery.EnumTipeItem.POB) Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TambahkanItemPO(ByVal vPrimary As Double, ByVal DefPrec As Integer)
        Dim rowDGV As DataGridViewRow
        Dim rowPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow
        Dim QtyBefore As Decimal = 0
        Dim NoPO As String = ""

        Try
            For Each rowPOB In Me.vTBLItemPOB.Rows
                If rowPOB.primary = vPrimary Then
                    'tambahkan baris ke dalam dgv, untuk begining balance item ready stock
                    Me.dgvData.Rows.Add()
                    rowDGV = Me.dgvData.Rows(Me.dgvData.Rows.Count - 1)
                    If rowDGV.Index = 0 Then
                        QtyBefore = 0
                    Else
                        If IsNumeric(Me.dgvData.Item(Me.ColBalance.Index, rowDGV.Index - 1).Value) = False Then
                            QtyBefore = 0
                        Else
                            QtyBefore = CDec(Me.dgvData.Item(Me.ColBalance.Index, rowDGV.Index - 1).Value)
                        End If
                    End If

                    rowDGV.Cells(Me.ColBalance.Index).Value = ObjCD.DigitGrouping(QtyBefore + rowPOB.qty, DefPrec)
                    If rowPOB.available_date = #1/1/2200# Then
                        rowDGV.Cells(Me.ColDate.Index).Value = "n/a"
                    Else
                        rowDGV.Cells(Me.ColDate.Index).Value = Format(rowPOB.available_date, "dd MMM yyyy")
                    End If
                    rowDGV.Cells(Me.ColDesc.Index).Value = "Purchase from " + ObjSupp.GetSupplierName(rowPOB.primary_supplier)
                    rowDGV.Cells(Me.ColEdit.Index).Value = "Edit"
                    rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvData.Rows.Count
                    rowDGV.Cells(Me.ColPrimary.Index).Value = rowPOB.primary
                    rowDGV.Cells(Me.ColPrimaryPriorityLevel.Index).Value = -1
                    rowDGV.Cells(Me.ColPriorityLevel.Index).Value = ""
                    rowDGV.Cells(Me.ColQtyIn.Index).Value = ObjCD.DigitGrouping(rowPOB.qty, DefPrec)
                    rowDGV.Cells(Me.ColQtyOut.Index).Value = "-"
                    rowDGV.Cells(Me.ColTipeItem.Index).Value = CInt(FrmTrans_MaterialPlanningDashboard_SF_DetailDelivery.EnumTipeItem.POB)
                    If rowPOB.tipe = ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.POB Then
                        NoPO = Me.ObjPO.GetNoTransaksiPO(rowPOB.primary_po)
                        If NoPO <> "" Then
                            rowDGV.Cells(Me.ColDocNo.Index).Value = NoPO
                        Else
                            rowDGV.Cells(Me.ColDocNo.Index).Value = rowPOB.po_number_rec
                        End If
                    Else
                        NoPO = Me.ObjPackingList.GetNoTransaksi(rowPOB.primary_packing_list)
                        If NoPO <> "" Then
                            rowDGV.Cells(Me.ColDocNo.Index).Value = NoPO
                        Else
                            rowDGV.Cells(Me.ColDocNo.Index).Value = rowPOB.packing_list_no_rec
                        End If
                    End If

                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungUlang()
        Dim rowItemPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow
        Dim IndexItemDelivery As Integer = 0
        Dim QtyPOReady As Decimal = 0
        Dim rowStock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockRow
        Dim rowItemSO_SF As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim IndexItemPOB As Integer = 0
        Dim Dummy As Decimal = 0
        Dim rowDGV As DataGridViewRow
        Dim DefPrec As Integer = 0
        Dim QtyReadyStock As Decimal = 0
        Dim QtyReadyStock1 As Decimal = 0
        Dim AdaBooked As Boolean = False

        Try
            Me.dgvData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvData.Rows.Clear()
            'jelaskan ke dalam dgv, kenapa tanggalnya bisa seperti hasil yang diberikan
            For Each rowItemPOB In Me.vTBLItemPOB.Rows
                'set parameter item pob yang akan digunakan pada perhitungan delivery date
                If rowItemPOB.primary_code_no = vPrimaryCodeNoTarget Then
                    rowItemPOB.available_date = rowItemPOB.promise_to_deliver + TimeSpan.FromDays(rowItemPOB.delivery_delay)
                    rowItemPOB.qty_used = 0
                End If
            Next
            'set parameter item sob yang akan digunakan pada perhitungan delivery date
            For Each rowItemSO_SF In Me.vTBLItemDelivery.Rows
                If rowItemSO_SF.primary_code_no = vPrimaryCodeNoTarget Then
                    rowItemSO_SF.qty_available = 0
                    rowItemSO_SF.estimate_delivery = #1/1/2200#
                    rowItemSO_SF.qty_ready_stock = 0
                End If
            Next
            'set parameter item stock untuk perhitungan ulang
            For Each rowStock In Me.vTBLItemStock.Rows
                rowStock.qty_used = 0
            Next

            'filter
            Me.vTBLItemStock.DefaultView.RowFilter = "[primary code no] = " + vPrimaryCodeNoTarget.ToString
            Me.vTBLItemPOB.DefaultView.RowFilter = "[primary code no] = " + vPrimaryCodeNoTarget.ToString
            Me.vTBLItemPOB.DefaultView.Sort = "[available date] ASC, [promise to deliver] ASC, [po date rec] ASC, [primary po] ASC"
            Me.vTBLItemDelivery.DefaultView.RowFilter = "[primary code no] = " + vPrimaryCodeNoTarget.ToString
            Me.vTBLItemDelivery.DefaultView.Sort = "[booked] DESC, [promise to deliver] ASC, [urutan priority level] ASC, [so date rec] ASC, [primary so] ASC"

            'set decimal
            If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                rowItemSO_SF = Me.vTBLItemDelivery.DefaultView(0).Row
                DefPrec = rowItemSO_SF._Decimal
            Else
                DefPrec = ObjProd.GetAngkaDecimal(vPrimaryCodeNoTarget)
            End If

            'kembalikan parameter ke nilai awal
            IndexItemDelivery = 0
            QtyPOReady = 0
            QtyReadyStock = 0

            'pasangkan dengan yang ready stock dahulu
            If Me.vTBLItemStock.DefaultView.Count > 0 Then
                rowStock = Me.vTBLItemStock.DefaultView(0).Row

                If rowStock.qty > 0 Then
                    'tambahkan baris ke dalam dgv, untuk begining balance item ready stock
                    Me.dgvData.Rows.Add()
                    rowDGV = Me.dgvData.Rows(Me.dgvData.Rows.Count - 1)
                    rowDGV.Cells(Me.ColBalance.Index).Value = ObjCD.DigitGrouping(rowStock.qty, DefPrec)
                    rowDGV.Cells(Me.ColDate.Index).Value = "-"
                    rowDGV.Cells(Me.ColDesc.Index).Value = "Stock On Hand"
                    rowDGV.Cells(Me.ColEdit.Index).Value = ""
                    rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvData.Rows.Count
                    rowDGV.Cells(Me.ColPrimary.Index).Value = rowStock.primary
                    rowDGV.Cells(Me.ColPrimaryPriorityLevel.Index).Value = -1
                    rowDGV.Cells(Me.ColPriorityLevel.Index).Value = ""
                    rowDGV.Cells(Me.ColQtyIn.Index).Value = "-"
                    rowDGV.Cells(Me.ColQtyOut.Index).Value = "-"
                    rowDGV.Cells(Me.ColTipeItem.Index).Value = CInt(FrmTrans_MaterialPlanningDashboard_SF_DetailDelivery.EnumTipeItem.Stock)
                End If

                While rowStock.qty - rowStock.qty_used > 0 And vTBLItemDelivery.DefaultView.Count > IndexItemDelivery
                    rowItemSO_SF = Me.vTBLItemDelivery.DefaultView(IndexItemDelivery).Row
                    If Me.vFromStockIn = False Then
                        If rowItemSO_SF.tipe <> CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req) Then
                            If rowStock.qty - rowStock.qty_used >= rowItemSO_SF.qty_sob Then
                                rowStock.qty_used += rowItemSO_SF.qty_sob
                                'karena ready stock, maka qty availabel = qty sob dan qty ready stock = qty sob
                                rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
                                rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                                'karena ready stock, maka tandai estimate delivery menjadi ASAP -> 1 jan 1950
                                rowItemSO_SF.estimate_delivery = #1/1/1950#
                                'tambahkan item delivery ke dalam dgv
                                Me.TambahkanItemDelivery(rowItemSO_SF.primary, rowItemSO_SF._Decimal)
                                'naikan index item delivery, cek item delivery selanjutnya
                                IndexItemDelivery += 1
                            Else
                                QtyReadyStock = rowStock.qty - rowStock.qty_used
                                rowStock.qty_used = rowStock.qty
                                Exit While
                            End If
                        Else
                            'naikan index item delivery, cek item delivery selanjutnya
                            IndexItemDelivery += 1
                        End If
                    Else
                        If rowStock.qty - rowStock.qty_used >= rowItemSO_SF.qty_sob Then
                            rowStock.qty_used += rowItemSO_SF.qty_sob
                            'karena ready stock, maka qty availabel = qty sob dan qty ready stock = qty sob
                            rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
                            rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                            'karena ready stock, maka tandai estimate delivery menjadi ASAP -> 1 jan 1950
                            rowItemSO_SF.estimate_delivery = #1/1/1950#
                            'tambahkan item delivery ke dalam dgv
                            Me.TambahkanItemDelivery(rowItemSO_SF.primary, rowItemSO_SF._Decimal)
                            'naikan index item delivery, cek item delivery selanjutnya
                            IndexItemDelivery += 1
                        Else
                            QtyReadyStock = rowStock.qty - rowStock.qty_used
                            rowStock.qty_used = rowStock.qty
                            Exit While
                        End If
                    End If

                End While
            End If

            'pasangkan dengan item po
            'editted by iwan: 2016-09-2016, karena bikin error fitri dok no. 60, tgl 25 agt 2016
            'IndexItemDelivery = 0
            If Me.vTBLItemPOB.DefaultView.Count > 0 Then
                IndexItemPOB = 0
                While IndexItemPOB < vTBLItemPOB.DefaultView.Count And IndexItemDelivery < vTBLItemDelivery.DefaultView.Count
                    rowItemPOB = Me.vTBLItemPOB.DefaultView(IndexItemPOB).Row
                    rowItemSO_SF = Me.vTBLItemDelivery.DefaultView(IndexItemDelivery).Row

                    'qty on hand hanya diperhitungkan jika item-nya sob
                    If rowItemSO_SF.tipe <> CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req) Then
                        QtyReadyStock1 = QtyReadyStock
                    Else
                        QtyReadyStock1 = 0
                    End If

                    'tambahkan item pob ke dalam dgv -> hanya jika belum ada saja
                    If ApakahSudahAda_Purchase(rowItemPOB.primary) = False Then
                        Me.TambahkanItemPO(rowItemPOB.primary, DefPrec)
                    End If

                    If QtyReadyStock1 + QtyPOReady + rowItemPOB.qty - rowItemPOB.qty_used >= rowItemSO_SF.qty_sob Then
                        'jika item yang tersedia lebih besar dari item delivery, maka 
                        'set delivery date + qty available delivery + pindah ke baris item delivery selanjutnya
                        rowItemSO_SF.estimate_delivery = rowItemPOB.available_date
                        rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
                        'majukan index item delivery
                        IndexItemDelivery += 1

                        'kurangkan qty on hand dan qty po ready
                        Dummy = rowItemSO_SF.qty_sob
                        If QtyReadyStock1 >= Dummy Then
                            rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                            QtyReadyStock1 -= Dummy
                            Dummy = 0
                        Else
                            Dummy -= QtyReadyStock1
                            rowItemSO_SF.qty_ready_stock = QtyReadyStock1
                            QtyReadyStock1 = 0

                            'jika ada masih ada sisa dummy, maka kurangkan dengan qty po ready (dari pob baris sebelumnya)
                            If QtyPOReady >= Dummy Then
                                QtyPOReady -= Dummy
                                Dummy = 0
                            Else
                                Dummy -= QtyPOReady
                                QtyPOReady = 0
                            End If
                        End If

                        'tambahkan qty used pada item po, untuk menanda jumlah yang sudah digunakan
                        rowItemPOB.qty_used += Dummy

                        'jika semua qty sudah habis dipakai, maka pindah ke baris item pob selanjutnya
                        If rowItemPOB.qty <= rowItemPOB.qty_used Then
                            IndexItemPOB += 1
                        End If

                        'tambahkan item delivery ke dalam dgv
                        Me.TambahkanItemDelivery(rowItemSO_SF.primary, DefPrec)
                    Else
                        'jika item delivery lebih besar dari qty po ready + qty po,ada 2 kemungkinan:
                        If IndexItemPOB + 1 < vTBLItemPOB.DefaultView.Count Then
                            '1. tambahkan qty po dari baris pob selanjut, karena dari pob, maka qty ready stock tidak di tambahkan
                            QtyPOReady += rowItemPOB.qty - rowItemPOB.qty_used
                            rowItemPOB.qty_used = rowItemPOB.qty
                            IndexItemPOB += 1
                        Else
                            '2. tidak ada baris pob selanjutnya, maka qty availble hanya sejumlah yang ada,
                            'estimate delivery tidak dikehui kapan, tapi available qty ada
                            rowItemSO_SF.qty_available = rowItemPOB.qty - rowItemPOB.qty_used + QtyPOReady + QtyReadyStock1
                            rowItemSO_SF.estimate_delivery = #1/1/2200#
                            'po ready di nol kan karena sudah habis
                            QtyPOReady = 0
                            'perhitungan ready stock
                            If QtyReadyStock1 >= rowItemSO_SF.qty_sob Then
                                rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                                QtyReadyStock1 -= rowItemSO_SF.qty_sob
                            Else
                                rowItemSO_SF.qty_ready_stock = QtyReadyStock1
                                QtyReadyStock1 = 0
                            End If
                            'tambahkan baris dgv
                            Me.TambahkanItemDelivery(rowItemSO_SF.primary, DefPrec)
                            'tambahan index item delivery, untuk sisanya
                            IndexItemDelivery += 1
                            'index item pob tetap dinaikan, untuk menghentikan looping perhitungan
                            IndexItemPOB += 1
                        End If
                    End If

                    'kembalikan posisi qty ready stock
                    If rowItemSO_SF.tipe <> CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req) Then
                        QtyReadyStock = QtyReadyStock1
                    End If
                End While
            End If
            'If vFromStockIn = True Then -> agak berbeda dengan form utama, sudah dihandle dibawah pada bagian sisa item sob yg belum dihandle
            '    If QtyReadyStock > 0 Then
            '        rowItemSO_SF = Me.vTBLItemDelivery.DefaultView(IndexItemDelivery).Row
            '        rowItemSO_SF.qty_available = QtyReadyStock
            '        rowItemSO_SF.qty_ready_stock = QtyReadyStock
            '        rowItemSO_SF.estimate_delivery = #1/1/2200#
            '        QtyReadyStock = 0
            '    End If
            'End If

            'jika masih ada sisa item pob yang belum di handle, maka tambahkan ke dgv
            IndexItemPOB += 1
            While IndexItemPOB < Me.vTBLItemPOB.DefaultView.Count
                rowItemPOB = Me.vTBLItemPOB.DefaultView(IndexItemPOB).Row
                Me.TambahkanItemPO(rowItemPOB.primary, DefPrec)
                IndexItemPOB += 1
            End While
            'jika masih ada sisa item sob yang belum di handle, maka tambahkan ke dgv
            While IndexItemDelivery < Me.vTBLItemDelivery.DefaultView.Count
                rowItemSO_SF = Me.vTBLItemDelivery.DefaultView(IndexItemDelivery).Row
                If Me.GetIndexItemDelivery(rowItemSO_SF.primary) = -1 Then
                    If rowItemSO_SF.qty_sob > rowItemSO_SF.qty_available Then
                        If QtyPOReady > 0 Then
                            rowItemSO_SF.qty_available += QtyPOReady
                            QtyPOReady = 0
                        End If
                        If QtyReadyStock > 0 Then
                            rowItemSO_SF.qty_available += QtyReadyStock
                            rowItemSO_SF.qty_ready_stock += QtyReadyStock
                            QtyReadyStock = 0
                        End If
                        Me.TambahkanItemDelivery(rowItemSO_SF.primary, DefPrec)
                    End If
                End If
                IndexItemDelivery += 1
            End While

            For Each rowDGV In Me.dgvData.Rows
                If rowDGV.Cells(Me.ColBookingNo.Index).Value <> "" Then
                    AdaBooked = True
                    Exit For
                End If
            Next
            If Me.CheckBox1.Checked = True Then
                Me.ColQtyPromiseToDeliver.Visible = True
                Me.ColPromiseToDeliver.Visible = True
                Me.ColDocNo.Visible = True
                If AdaBooked = True Then
                    Me.ColBookingNo.Visible = True
                    Me.ColBookQty.Visible = True
                Else
                    Me.ColBookingNo.Visible = False
                    Me.ColBookQty.Visible = False
                End If
                Me.ColCreatedBy.Visible = True
            Else
                Me.ColQtyPromiseToDeliver.Visible = False
                Me.ColPromiseToDeliver.Visible = False
                Me.ColDocNo.Visible = False
                Me.ColBookingNo.Visible = False
                Me.ColBookQty.Visible = False
                Me.ColCreatedBy.Visible = False
            End If

            Me.dgvData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemDelivery(ByVal vPrimary As Double) As Integer
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Integer = 0

        Try
            For Each rowDGV In Me.dgvData.Rows
                If rowDGV.Cells(Me.ColTipeItem.Index).Value = CInt(FrmTrans_MaterialPlanningDashboard_SF_DetailDelivery.EnumTipeItem.SOB) _
                And rowDGV.Cells(Me.ColPrimary.Index).Value = vPrimary Then
                    Return RetVal
                End If
                RetVal += 1
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Public Function GetTanggalTelatPertama() As Date
        Dim rowDel As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow

        Try
            Me.vTBLItemDelivery.DefaultView.RowFilter = ""
            Me.vTBLItemDelivery.DefaultView.Sort = "promise to deliver"

            For Each dv As Data.DataRowView In Me.vTBLItemDelivery.DefaultView
                rowDel = dv.Row
                If rowDel.qty_available > 0 Then
                    If rowDel.qty_ready_stock >= rowDel.qty_available Then
                        'ga telat
                    Else
                        'partial atau waiting
                        If rowDel.promise_to_deliver <> #1/1/2200# And rowDel.promise_to_deliver <> #1/1/1950# Then
                            Return rowDel.promise_to_deliver
                        End If
                    End If
                Else
                    If rowDel.promise_to_deliver <> #1/1/2200# And rowDel.promise_to_deliver <> #1/1/1950# Then
                        Return rowDel.promise_to_deliver
                    End If
                End If
            Next

            Return #1/1/1950#
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub FrmTrans_MaterialPlanningDashboard_SFDetailDelivery_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ListDelPrio As New List(Of String)

        Try
            'tambahkan priority level
            Me.ObjPriorityLevel.GetListDescPriorityLevel(ListDelPrio)
            Me.ColPriorityLevel.Items.Clear()
            Me.ColPriorityLevel.Items.Add("")
            For Each Priority As String In ListDelPrio
                Me.ColPriorityLevel.Items.Add(Priority)
            Next
            Me.HitungUlang()
            Me.dgvData.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvData.AllowUserToAddRows = False
            Me.dgvData.AllowUserToDeleteRows = False
            If Me.vReadOnly = True Then
                Me.dgvData.ReadOnly = True
                Me.ColEdit.Visible = False
            End If

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_SF_Override_SalesETD") = False _
            And Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_SF_Override_PurchaseETD") = False Then
                Me.ColEdit.Visible = False
            End If

            If Me.ObjProd.ApakahGoods(vPrimaryCodeNoTarget) = False Then
                TipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Services
            Else
                TipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
            End If
            If TipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Services Then
                Me.ColBalance.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListDelPrio = Nothing
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Function ApakahDataMasihSamaDgDB_POB(ByVal vPrimary As Double) As Boolean
        Dim Identifier As String = ""
        Dim rowPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow
        Dim DTA1 As New DseTransaksiShippingForecastTableAdapters.tbltrans_po_itemdelivery1TableAdapter
        Dim TBL1 As New DseTransaksiShippingForecast.tbltrans_po_itemdelivery1DataTable
        Dim row1 As DseTransaksiShippingForecast.tbltrans_po_itemdelivery1Row
        Dim DummyFilter, DummySort As String
        Dim Ketemu As Boolean = False

        Try
            DummySort = Me.vTBLItemPOB.DefaultView.Sort
            DummyFilter = Me.vTBLItemPOB.DefaultView.RowFilter
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            Me.vTBLItemPOB.DefaultView.RowFilter = "primary = " + vPrimary.ToString
            If Me.vTBLItemPOB.DefaultView.Count > 0 Then
                rowPOB = Me.vTBLItemPOB.DefaultView(0).Row
                Identifier = rowPOB.identifier
            Else
                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan po-nya tidak dalam posisi removed
            If Me.ObjPO.ApakahRemoved(rowPOB.primary_po) = True Then
                MsgBox("The po associated to this is removed, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan identifiernya masih ada 
            If Me.ObjPOItem.GetJmlProductByIdentifier(Identifier) < 0 Then
                MsgBox("We cannot find associated po item for this data, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan primary code no-nya masih sama
            If Me.ObjPOItem.GetPrimaryProductByIdentifier(Identifier) <> Me.vPrimaryCodeNoTarget Then
                MsgBox("The associated po item for this data is using different code no, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If rowPOB.tipe = ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.POB Then
                'pastikan setiap kolom dalam item deliverynya juga masih sama, yang dimaksud sama adalah:
                '1. jumlah item delivery qty sama
                '2. jika diurutkan dari atas, delivery qty, delay dan delivery datenya harus sama
                DTA1.Fill(TBL1, Identifier)
                Me.vTBLItemPOB.DefaultView.RowFilter = "identifier ='" + Identifier + "'"
                Me.vTBLItemPOB.DefaultView.Sort = "[promise to deliver] ASC, qty ASC"
                If TBL1.Rows.Count <> 0 Then '-> jika item po tersebut tidak delivery cont, maka tidak perlu mencocokan jumlah item, bagian batasan ini lewat
                    If Me.vTBLItemPOB.DefaultView.Count <> TBL1.Rows.Count Then
                        MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                        Return False
                    Else
                        For ij As Integer = 0 To Me.vTBLItemPOB.DefaultView.Count - 1
                            rowPOB = Me.vTBLItemPOB.DefaultView(ij).Row
                            row1 = TBL1.Rows(ij)
                            Ketemu = True
                            If rowPOB.delivery_delay <> row1.delay Then
                                MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                            If rowPOB.promise_to_deliver <> row1.delivery_date Then
                                MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                            If rowPOB.qty <> row1.delivery_qty Then
                                MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                            Exit For
                        Next
                        If Ketemu = False Then
                            MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            Else
                'pastikan packing listnya masih ada
                If Me.ObjPackingList.GetNoTransaksi(rowPOB.primary_packing_list) = "" Then
                    MsgBox("We cannot find the specified packing list, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                'pastikan ada kombinasi code no + po ybs pada packing list tersebut
                If Me.ObjPackingList.GetJml(rowPOB.primary_packing_list, rowPOB.primary_code_no, rowPOB.primary_po) = 0 Then
                    MsgBox("We cannot find the following combination: Code No: " + Me.ObjProd.GetProductCode(rowPOB.primary_code_no) + " - PO Number: " + Me.ObjPO.GetNoTransaksiPO(rowPOB.primary_po) + ", please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                'pastikan estimate delivery pada packing listnya sama
                If rowPOB.promise_to_deliver.Date <> Me.ObjPackingList.Get_EstimateArrival(rowPOB.primary_packing_list).Date Then
                    MsgBox("The estimate delivery for the associated packing list does not match, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                'pastikan packing listnya belum di stock in
                If Me.ObjPackingList.ApakahSudahStockIn(rowPOB.primary_packing_list) = True Then
                    MsgBox("The associated packing list already received, cannot edit this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Me.vTBLItemPOB.DefaultView.Sort = DummySort
            Me.vTBLItemPOB.DefaultView.RowFilter = DummyFilter

            'sampa sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Function

    Public Function ApakahRemoved(ByVal vPrimaryPO As Double) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
        Dim Retval As Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetIsRemoved(vPrimaryPO)

            If Retval.HasValue = False Then
                Return False
            Else
                If Retval.Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function ApakahDataMasihSamaDgDB_POB_EditLink(ByVal vPrimary As Double) As Boolean
        Dim Identifier As String = ""
        Dim rowPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow
        Dim DTA1 As New DseTransaksiShippingForecastTableAdapters.tbltrans_po_itemdelivery1TableAdapter
        Dim TBL1 As New DseTransaksiShippingForecast.tbltrans_po_itemdelivery1DataTable
        Dim row1 As DseTransaksiShippingForecast.tbltrans_po_itemdelivery1Row
        Dim DummyFilter, DummySort As String
        Dim Ketemu As Boolean = False

        Try
            DummySort = Me.vTBLItemPOB.DefaultView.Sort
            DummyFilter = Me.vTBLItemPOB.DefaultView.RowFilter
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            Me.vTBLItemPOB.DefaultView.RowFilter = "primary = " + vPrimary.ToString
            If Me.vTBLItemPOB.DefaultView.Count > 0 Then
                rowPOB = Me.vTBLItemPOB.DefaultView(0).Row
                Identifier = rowPOB.identifier
            Else
                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan po-nya tidak dalam posisi removed
            If Me.ObjPO.ApakahRemoved(rowPOB.primary_po) = True Then
                MsgBox("The po associated to this is removed, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan identifiernya masih ada 
            If Me.ObjPOItem.GetJmlProductByIdentifier(Identifier) < 0 Then
                MsgBox("We cannot find associated po item for this data, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan primary code no-nya masih sama
            If Me.ObjPOItem.GetPrimaryProductByIdentifier(Identifier) <> Me.vPrimaryCodeNoTarget Then
                MsgBox("The associated po item for this data is using different code no, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan pob > 0 , sebab jika pob > 0, maka akan bikin masalah ketika hitung ulang, karena pob = 0, maka item po tidak akan muncul dalam hitung ulang
            If Me.ObjPOItem.GetOrderBalanceByIdentifier1(rowPOB.identifier) = 0 Then
                MsgBox("You cannot edit this data because current order balance data is 0.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If rowPOB.tipe = ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.POB Then
                'pastikan setiap kolom dalam item deliverynya juga masih sama, yang dimaksud sama adalah:
                '1. jumlah item delivery qty sama
                '2. jika diurutkan dari atas, delivery qty, delay dan delivery datenya harus sama
                DTA1.Fill(TBL1, Identifier)
                Me.vTBLItemPOB.DefaultView.RowFilter = "identifier ='" + Identifier + "'"
                Me.vTBLItemPOB.DefaultView.Sort = "[promise to deliver] ASC, qty ASC"
                If TBL1.Rows.Count <> 0 Then '-> jika item po tersebut tidak delivery cont, maka tidak perlu mencocokan jumlah item, bagian batasan ini lewat
                    If Me.vTBLItemPOB.DefaultView.Count <> TBL1.Rows.Count Then
                        MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                        Return False
                    Else
                        For ij As Integer = 0 To Me.vTBLItemPOB.DefaultView.Count - 1
                            rowPOB = Me.vTBLItemPOB.DefaultView(ij).Row
                            row1 = TBL1.Rows(ij)
                            Ketemu = True
                            If rowPOB.delivery_delay <> row1.delay Then
                                MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                            If rowPOB.promise_to_deliver <> row1.delivery_date Then
                                MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                            If rowPOB.qty <> row1.delivery_qty Then
                                MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                            Exit For
                        Next
                        If Ketemu = False Then
                            MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            Else
                MsgBox("This item referred by packing list you cannot edit this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Me.vTBLItemPOB.DefaultView.Sort = DummySort
            Me.vTBLItemPOB.DefaultView.RowFilter = DummyFilter

            'sampa sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Function

    Private Function ApakahDataMasihSamaDgnDB_SOB(ByVal vPrimary As Double) As Boolean
        Dim Identifier As String = ""
        Dim rowSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim DTA1 As New DseTransaksiShippingForecastTableAdapters.tbltrans_so_itemdelivery1TableAdapter
        Dim TBL1 As New DseTransaksiShippingForecast.tbltrans_so_itemdelivery1DataTable
        Dim row1 As DseTransaksiShippingForecast.tbltrans_so_itemdelivery1Row
        Dim DummyFilter, DummySort As String

        Try
            DummySort = Me.vTBLItemDelivery.DefaultView.Sort
            DummyFilter = Me.vTBLItemDelivery.DefaultView.RowFilter

            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            Me.vTBLItemDelivery.DefaultView.RowFilter = "primary = " + vPrimary.ToString
            If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                rowSOB = Me.vTBLItemDelivery.DefaultView(0).Row
                Identifier = rowSOB.identifier
            Else
                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan identifiernya masih ada 
            If Me.ObjSOItem.GetJmlProductByIdentifier(Identifier) < 0 Then
                MsgBox("We cannot find associated so item for this data, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan primary code no-nya masih sama
            If Me.ObjSOItem.GetPrimaryProductByIdentifier(Identifier) <> Me.vPrimaryCodeNoTarget Then
                MsgBox("The associated so item for this data is using different code no, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan delivery date controlled, klo tidak maka tidak bisa edit data priority level,
            'karena bikin pusing aj codingnya, belum ketemu perlakukannya yang baik
            If Me.ObjSOItem.ApakahDeliveryDateControl(Identifier) = False Then
                MsgBox("The associated so item for this data is not using delivery date control, please use edit link if you want to change this.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan setiap kolom dalam item deliverynya juga masih sama, yang dimaksud sama adalah:
            '1. jika diurutkan dari bawah, jumlah delivery qty dan delivery datenya sama -> untuk item terakhir ada kemungkinan tidak sama karena
            'pada perhitungan shipping forecast jumlahnya hanya diambil sesuai SOB-tidak diambil jumlah so penuh
            DTA1.Fill(TBL1, Identifier)
            If TBL1.Rows.Count = 0 Then
                MsgBox("The associated so item for this data is not using delivery date control, please use edit link if you want to change this.", _
                MsgBoxStyle.Information, "Information")
                Return False
            Else
                Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier ='" + Identifier + "'"
                Me.vTBLItemDelivery.DefaultView.Sort = "[promise to deliver] DESC, [qty sob] DESC"
                For ij As Integer = 0 To Me.vTBLItemDelivery.DefaultView.Count - 1
                    rowSOB = Me.vTBLItemDelivery.DefaultView(ij).Row
                    row1 = TBL1.Rows(ij)
                    If rowSOB.primary_priority_level <> row1.delivery_priority Then
                        MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    If rowSOB.promise_to_deliver <> row1.delivery_date Then
                        MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    'jika baris terakhir tidak perlu cek masalah ini
                    If ij <> Me.vTBLItemDelivery.DefaultView.Count - 1 Then
                        If rowSOB.qty_sob <> row1.delivery_qty Then
                            MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                Next
            End If

            Me.vTBLItemDelivery.DefaultView.Sort = DummySort
            Me.vTBLItemDelivery.DefaultView.RowFilter = DummyFilter

            'sampa sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Function

    Private Function ApakahDataMasihSamaDgnDB_PurcReq(ByVal vPrimary As Double) As Boolean
        Dim Identifier As String = ""
        Dim rowSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim DTA1 As New DseTransaksiShippingForecastTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
        Dim TBL1 As New DseTransaksiShippingForecast.tbltrans_purchase_req_itemdeliveryDataTable
        Dim row1 As DseTransaksiShippingForecast.tbltrans_purchase_req_itemdeliveryRow
        Dim DummyFilter, DummySort As String

        Try
            DummySort = Me.vTBLItemDelivery.DefaultView.Sort
            DummyFilter = Me.vTBLItemDelivery.DefaultView.RowFilter

            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            Me.vTBLItemDelivery.DefaultView.RowFilter = "primary = " + vPrimary.ToString
            If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                rowSOB = Me.vTBLItemDelivery.DefaultView(0).Row
                Identifier = rowSOB.identifier
            Else
                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan identifiernya masih ada 
            If Me.ObjPurcReq.GetJmlByPrimaryCode(rowSOB.primary_so, rowSOB.primary_code_no) <= 0 Then
                MsgBox("We cannot find associated purchase request item for this data, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan delivery date controlled, klo tidak maka tidak bisa edit data priority level,
            'karena bikin pusing aj codingnya, belum ketemu perlakukannya yang baik
            If Me.ObjPurcReq.ApakahDeliveryDateControl(rowSOB.primary_so, rowSOB.primary_code_no) = False Then
                MsgBox("The associated purchase request item for this data is not using delivery date control, please use edit link if you want to change this.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan setiap kolom dalam item deliverynya juga masih sama, yang dimaksud sama adalah:
            '1. jika diurutkan dari bawah, jumlah delivery qty dan delivery datenya sama -> untuk item terakhir ada kemungkinan tidak sama karena
            'pada perhitungan shipping forecast jumlahnya hanya diambil sesuai SOB-tidak diambil jumlah so penuh
            DTA1.FillBy_Identifier(TBL1, Identifier)
            If TBL1.Rows.Count = 0 Then
                MsgBox("The associated so item for this data is not using delivery date control, please use edit link if you want to change this.", _
                MsgBoxStyle.Information, "Information")
                Return False
            Else
                Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier ='" + Identifier + "'"
                Me.vTBLItemDelivery.DefaultView.Sort = "[promise to deliver] DESC, [qty sob] DESC"
                For ij As Integer = 0 To Me.vTBLItemDelivery.DefaultView.Count - 1
                    rowSOB = Me.vTBLItemDelivery.DefaultView(ij).Row
                    row1 = TBL1.Rows(ij)
                    If rowSOB.primary_priority_level <> row1.delivery_priority Then
                        MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    If rowSOB.promise_to_deliver <> row1.delivery_date Then
                        MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    'jika baris terakhir tidak perlu cek masalah ini
                    If ij <> Me.vTBLItemDelivery.DefaultView.Count - 1 Then
                        If rowSOB.qty_sob <> row1.delivery_qty Then
                            MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                Next
            End If

            Me.vTBLItemDelivery.DefaultView.Sort = DummySort
            Me.vTBLItemDelivery.DefaultView.RowFilter = DummyFilter

            'sampa sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Function

    Private Function ApakahDataMasihSamaDgnDB_PurchReq_EditLink(ByVal vPrimary As Double) As Boolean
        Dim Identifier As String = ""
        Dim rowSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim DTA1 As New DseTransaksiShippingForecastTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
        Dim TBL1 As New DseTransaksiShippingForecast.tbltrans_purchase_req_itemdeliveryDataTable
        Dim row1 As DseTransaksiShippingForecast.tbltrans_purchase_req_itemdeliveryRow
        Dim DummyFilter, DummySort As String

        Try
            DummySort = Me.vTBLItemDelivery.DefaultView.Sort
            DummyFilter = Me.vTBLItemDelivery.DefaultView.RowFilter

            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            Me.vTBLItemDelivery.DefaultView.RowFilter = "primary = " + vPrimary.ToString
            If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                rowSOB = Me.vTBLItemDelivery.DefaultView(0).Row
                Identifier = rowSOB.identifier
            Else
                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan identifiernya masih ada 
            If Me.ObjPurcReq.GetJmlByPrimaryCode(rowSOB.primary_so, rowSOB.primary_code_no) <= 0 Then
                MsgBox("We cannot find associated purchase request item for this data, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan sob > 0 , sebab jika sob > 0, maka akan bikin masalah ketika hitung ulang, karena sob = 0, maka item po tidak akan muncul dalam hitung ulang
            If Me.ObjPurcReq.GetOrderBalance_Only(rowSOB.primary_so, rowSOB.primary_code_no) = 0 Then
                MsgBox("You cannot edit this data because current order balance data is 0.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan setiap kolom dalam item deliverynya juga masih sama, yang dimaksud sama adalah:
            '1. jika diurutkan dari bawah, jumlah delivery qty dan delivery datenya sama -> untuk item terakhir ada kemungkinan tidak sama karena
            'pada perhitungan shipping forecast jumlahnya hanya diambil sesuai SOB-tidak diambil jumlah so penuh
            DTA1.FillBy_Identifier(TBL1, Identifier)
            If TBL1.Rows.Count = 0 Then
                'tidak perlu pengecekan, klo tidak delivery date control
            Else
                Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier ='" + Identifier + "'"
                Me.vTBLItemDelivery.DefaultView.Sort = "[promise to deliver] DESC, [qty sob] DESC"
                For ij As Integer = 0 To Me.vTBLItemDelivery.DefaultView.Count - 1
                    rowSOB = Me.vTBLItemDelivery.DefaultView(ij).Row
                    row1 = TBL1.Rows(ij)
                    If rowSOB.primary_priority_level <> row1.delivery_priority Then
                        MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    If rowSOB.promise_to_deliver <> row1.delivery_date Then
                        MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    'jika baris terakhir tidak perlu cek masalah ini
                    If ij <> Me.vTBLItemDelivery.DefaultView.Count - 1 Then
                        If rowSOB.qty_sob <> row1.delivery_qty Then
                            MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                Next
            End If

            Me.vTBLItemDelivery.DefaultView.Sort = DummySort
            Me.vTBLItemDelivery.DefaultView.RowFilter = DummyFilter

            'sampa sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Function

    Private Function ApakahDataMasihSamaDgnDB_SOB_EditLink(ByVal vPrimary As Double) As Boolean
        Dim Identifier As String = ""
        Dim rowSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim DTA1 As New DseTransaksiShippingForecastTableAdapters.tbltrans_so_itemdelivery1TableAdapter
        Dim TBL1 As New DseTransaksiShippingForecast.tbltrans_so_itemdelivery1DataTable
        Dim row1 As DseTransaksiShippingForecast.tbltrans_so_itemdelivery1Row
        Dim DummyFilter, DummySort As String

        Try
            DummySort = Me.vTBLItemDelivery.DefaultView.Sort
            DummyFilter = Me.vTBLItemDelivery.DefaultView.RowFilter

            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            Me.vTBLItemDelivery.DefaultView.RowFilter = "primary = " + vPrimary.ToString
            If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                rowSOB = Me.vTBLItemDelivery.DefaultView(0).Row
                Identifier = rowSOB.identifier
            Else
                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan identifiernya masih ada 
            If Me.ObjSOItem.GetJmlProductByIdentifier(Identifier) < 0 Then
                MsgBox("We cannot find associated so item for this data, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan primary code no-nya masih sama
            If Me.ObjSOItem.GetPrimaryProductByIdentifier(Identifier) <> Me.vPrimaryCodeNoTarget Then
                MsgBox("The associated so item for this data is using different code no, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan sob > 0 , sebab jika sob > 0, maka akan bikin masalah ketika hitung ulang, karena sob = 0, maka item po tidak akan muncul dalam hitung ulang
            If Me.ObjSOItem.GetOrderBalanceByIdentifier(rowSOB.identifier) = 0 Then
                MsgBox("You cannot edit this data because current order balance data is 0.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan setiap kolom dalam item deliverynya juga masih sama, yang dimaksud sama adalah:
            '1. jika diurutkan dari bawah, jumlah delivery qty dan delivery datenya sama -> untuk item terakhir ada kemungkinan tidak sama karena
            'pada perhitungan shipping forecast jumlahnya hanya diambil sesuai SOB-tidak diambil jumlah so penuh
            DTA1.Fill(TBL1, Identifier)
            If TBL1.Rows.Count = 0 Then
                'tidak perlu pengecekan, klo tidak delivery date control
            Else
                Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier ='" + Identifier + "'"
                Me.vTBLItemDelivery.DefaultView.Sort = "[promise to deliver] DESC, [qty sob] DESC"
                For ij As Integer = 0 To Me.vTBLItemDelivery.DefaultView.Count - 1
                    rowSOB = Me.vTBLItemDelivery.DefaultView(ij).Row
                    row1 = TBL1.Rows(ij)
                    If rowSOB.primary_priority_level <> row1.delivery_priority Then
                        MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    If rowSOB.promise_to_deliver <> row1.delivery_date Then
                        MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    'jika baris terakhir tidak perlu cek masalah ini
                    If ij <> Me.vTBLItemDelivery.DefaultView.Count - 1 Then
                        If rowSOB.qty_sob <> row1.delivery_qty Then
                            MsgBox("Delivery item count is not valid, please recalculate shipping forecast.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                Next
            End If

            Me.vTBLItemDelivery.DefaultView.Sort = DummySort
            Me.vTBLItemDelivery.DefaultView.RowFilter = DummyFilter

            'sampa sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Function

    Private Sub dgvData_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvData.CellBeginEdit
        Try
            If Me.vReadOnly = True Then
                Exit Sub
            End If
            'yang boleh di edit adalah:
            '1. tanggal delivery date untuk po item dan so item
            '2. priority level untuk so item
            ', dengan catatan: item yang bersangkutan memang belum dikirimkan sehingga valid untuk melakukan edit data
            'on second thouht, hal ini diperboleh saja, karena pertimbangan misalnya nanti klo ada laporan evaluasi supplier
            'biasanya ini di ubah sehingga laporannya tidak terlalu parah.
            'If e.ColumnIndex = Me.ColDate.Index Or e.ColumnIndex = Me.ColPriorityLevel.Index Then
            'pastikan datanya masih cocok dengan database, jika datanya sudah tidak cocok, maka tidak bisa diedit
            'artinya sudah ada user lain yang melakukan perubahan data ini
            If Me.dgvData.Item(Me.ColTipeItem.Index, e.RowIndex).Value = CInt(FrmTrans_MaterialPlanningDashboard_SF_DetailDelivery.EnumTipeItem.POB) Then
                If e.ColumnIndex = Me.ColDate.Index Then
                    If Me.ApakahDataMasihSamaDgDB_POB(Me.dgvData.Item(Me.ColPrimary.Index, e.RowIndex).Value) = True Then
                        Me.Buffer = Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    e.Cancel = True
                End If
            ElseIf Me.dgvData.Item(Me.ColTipeItem.Index, e.RowIndex).Value = CInt(FrmTrans_MaterialPlanningDashboard_SF_DetailDelivery.EnumTipeItem.SOB) Then
                If e.ColumnIndex = Me.ColPriorityLevel.Index Then
                    If Me.dgvData.Item(Me.ColTipeInt.Index, e.RowIndex).Value = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB Then
                        If Me.ApakahDataMasihSamaDgnDB_SOB(Me.dgvData.Item(Me.ColPrimary.Index, e.RowIndex).Value) = True Then
                            Me.Buffer = Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    Else
                        If Me.ApakahDataMasihSamaDgnDB_PurcReq(Me.dgvData.Item(Me.ColPrimary.Index, e.RowIndex).Value) = True Then
                            Me.Buffer = Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvData_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvData.CellEndEdit
        Try
            If IsNothing(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If

            If FrmTrans_MaterialPlanningDashboard_SF_DetailDelivery.EnumTipeItem.SOB = CInt(Me.dgvData.Item(Me.ColTipeItem.Index, e.RowIndex).Value) Then
                If Me.dgvData.Item(Me.ColTipeInt.Index, e.RowIndex).Value = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB Then
                    Me.EndEditPriorityLevel(sender, e)
                Else
                    Me.EndEditPriorityLevel_PurchaseReq(sender, e)
                End If
            ElseIf FrmTrans_MaterialPlanningDashboard_SF_DetailDelivery.EnumTipeItem.POB = CInt(Me.dgvData.Item(Me.ColTipeItem.Index, e.RowIndex).Value) Then
                Me.EndEdit_PODelDate(sender, e)
            Else
                Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEdit_PackingListDate(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, ByVal rowPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow)
        Try
            'ubah ke dalam item pob
            rowPOB.promise_to_deliver = CDate(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) - TimeSpan.FromDays(rowPOB.delivery_delay)
            rowPOB.available_date = CDate(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value)

            'update ke dalam database -> set estimate delivery untuk packing list ybs
            Me.ObjPackingList.UpdateEstimateDelivery(rowPOB.primary_packing_list, rowPOB.promise_to_deliver)

            'hitung ulang + posisikan ulang
            Me.BackgroundWorker1.RunWorkerAsync()
            Me.DataChanged = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEdit_PODelDate(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim rowPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow
        Dim Identifier As String = ""
        Dim DummyDelay As Integer = 0

        Try
            'tidak boleh string kosong
            If Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value.ToString = "" Then
                Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If
            'harus tanggal
            If IsDate(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                MsgBox("Please secify a valid date.", MsgBoxStyle.Information, "Information")
                Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If
            'kalo sama dengan buffer maka tidak di proses
            If IsDate(Me.Buffer) = True Then
                If CDate(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) = CDate(Me.Buffer) Then
                    Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
            Else
                If Me.Buffer.ToLower = Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToLower Then
                    Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
            End If

            'karena asumsi pengecekanya sama(sudah dicek waktu cell begin edit), maka ambil saja dari 
            Me.vTBLItemPOB.DefaultView.RowFilter = "primary = " + Me.dgvData.Item(Me.ColPrimary.Index, e.RowIndex).Value.ToString
            If Me.vTBLItemPOB.DefaultView.Count > 0 Then
                rowPOB = Me.vTBLItemPOB.DefaultView(0).Row
                Identifier = rowPOB.identifier
            Else
                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If

            If rowPOB.tipe = ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.GIT Then
                Me.EndEdit_PackingListDate(sender, e, rowPOB)
                Exit Sub
            End If

            'tabel pob item yang digunakan perhitungan, tidak perlu ambil dari db
            Me.vTBLItemPOB.DefaultView.RowFilter = "identifier = '" + Identifier + "'"
            Me.vTBLItemPOB.DefaultView.Sort = "[promise to deliver] ASC"
            If Me.vTBLItemPOB.DefaultView.Count > 0 Then
                If Me.vTBLItemPOB.DefaultView.Count > 1 Then
                    'delivery date tidak boleh sama atau kurang dari item sebelumnya, 
                    For ij As Integer = 1 To Me.vTBLItemPOB.DefaultView.Count - 1
                        rowPOB = Me.vTBLItemPOB.DefaultView(ij).Row
                        DummyDelay = rowPOB.delivery_delay
                        If rowPOB.primary = Me.dgvData.Item(Me.ColPrimary.Index, e.RowIndex).Value Then
                            rowPOB = Me.vTBLItemPOB.DefaultView(ij - 1).Row
                            If (CDate(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value).Date - TimeSpan.FromDays(DummyDelay)).Date <= (rowPOB.promise_to_deliver).Date Then
                                MsgBox("Minimum delivery date is: " + Format((rowPOB.promise_to_deliver + TimeSpan.FromDays(DummyDelay + 1)).Date, "dd MMM yyyy"), _
                                MsgBoxStyle.Information, "Information")
                                Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If
                    Next
                    'delivery date tidak boleh lebih besar atau sama dari item setelah
                    For ij As Integer = 0 To Me.vTBLItemPOB.DefaultView.Count - 2
                        rowPOB = Me.vTBLItemPOB.DefaultView(ij).Row
                        DummyDelay = rowPOB.delivery_delay
                        If rowPOB.primary = Me.dgvData.Item(Me.ColPrimary.Index, e.RowIndex).Value Then
                            rowPOB = Me.vTBLItemPOB.DefaultView(ij + 1).Row
                            If (CDate(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value).Date - TimeSpan.FromDays(DummyDelay)).Date >= (rowPOB.promise_to_deliver).Date Then
                                MsgBox("Maximum delivery date is: " + Format((rowPOB.promise_to_deliver + TimeSpan.FromDays(DummyDelay - 1)).Date, "dd MMM yyyy"), _
                                MsgBoxStyle.Information, "Information")
                                Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If
                    Next
                End If

                'tanggalnya minimal sama dengan tanggal po, tidak boleh sama dengan tanggal po
                If Me.ObjPO.GetTanggalPO(rowPOB.primary_po) > CDate(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    If MsgBox("Delivery date is smaller than po date, continue?", _
                    MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
                'update ke dalam po item
                For ij As Integer = 0 To Me.vTBLItemPOB.DefaultView.Count - 1
                    rowPOB = Me.vTBLItemPOB.DefaultView(ij).Row
                    If rowPOB.primary = Me.dgvData.Item(Me.ColPrimary.Index, e.RowIndex).Value Then
                        'ubah ke dalam item pob
                        rowPOB.promise_to_deliver = CDate(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) - TimeSpan.FromDays(rowPOB.delivery_delay)
                        rowPOB.available_date = CDate(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value)

                        'update ke dalam database
                        If Me.ObjPOItem.ApakahDeliveryDateCont(Identifier) = True Then
                            '1. jika ada item delivery date, maka update item ybs saja
                            Me.ObjPOItem.UpdateDeliveryDate(Identifier, CDate(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) - TimeSpan.FromDays(rowPOB.delivery_delay), _
                            CDate(Me.Buffer) - TimeSpan.FromDays(rowPOB.delivery_delay))
                        Else
                            '2. jika tidak ada item delivery, maka:
                            '1. ubah item menjadi delivery date controlled + masukan item delivery date dengan tanggal seperti tanggal yang di spesifikasikan 
                            ' + qty = qty item po + delay = 0
                            Me.ObjPOItem.SetDeliveryDateControlled(Identifier, CDate(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value))
                            '2. ubah timestamp untuk po tsb
                            Me.ObjPOItem.UpdateTimestamp(rowPOB.primary_po, Now)
                        End If

                        'hitung ulang + posisikan ulang
                        Me.BackgroundWorker1.RunWorkerAsync()
                        Me.DataChanged = True

                        'exit sub 
                        Exit Sub
                    End If
                Next
            Else
                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub UpdateTimestamp(ByVal vPrimaryPO As Double, ByVal vTimestamp As Date)
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetTimestamp(vTimestamp, vPrimaryPO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub EndEditPriorityLevel(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_so_itemdelivery1TableAdapter
        Dim TBL As New DseTransaksiShippingForecast.tbltrans_so_itemdelivery1DataTable
        Dim Identifier As String = ""
        Dim rowSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim rowAcuan, rowDummy As DseTransaksiShippingForecast.tbltrans_so_itemdelivery1Row
        Dim DTA1 As New DseTransaksiShippingForecastTableAdapters.tbltrans_so_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            'tidak boleh string kosong
            If Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If

            'kalo sama dengan buffer tidak di proses
            If Me.Buffer.ToLower = Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToLower Then
                Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If

            'ambil data identifier
            Me.vTBLItemDelivery.DefaultView.RowFilter = "primary = " + Me.dgvData.Item(Me.ColPrimary.Index, e.RowIndex).Value.ToString
            If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                rowSOB = Me.vTBLItemDelivery.DefaultView(0).Row
                Identifier = rowSOB.identifier

                'dilewat sementara, karena memungkinkan untuk terbalik, priority levelnya, misalnya order yang kedua high priority
                'pastikan priority level selalu lebih kecil atau sama dengan index priority level di atas nya
                'hanya cek jika bukan baris pertama
                'DTA.FillBy_ASC(TBL, Identifier)
                'For ij As Integer = 0 To TBL.Rows.Count - 1
                '    rowDummy = TBL.Rows(ij)
                '    If rowDummy.delivery_date = rowSOB.promise_to_deliver Then
                '        If ij <> 0 Then
                '            rowDummy1 = TBL.Rows(ij - 1)
                '            rowAcuan = TBL.Rows(ij)

                '            If Me.ObjPriorityLevel.GetUrutan(rowDummy1.delivery_priority) > _
                '            Me.ObjPriorityLevel.GetUrutan(Me.ObjPriorityLevel.GetPrimary_ByDesc(Me.dgvData.Item(Me.ColPriorityLevel.Index, e.RowIndex).Value)) Then
                '                MsgBox("Maximum priority is " + Me.ObjPriorityLevel.GetDesc(rowDummy1.delivery_priority), MsgBoxStyle.Information, "Information")
                '                Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                '                Exit Sub
                '            End If
                '        End If
                '    End If
                'Next

                'pastikan priority level selalu lebih besar dari bawahnya
                'hanya cek jika bukan baris terakhir
                DTA.FillBy_ASC(TBL, Identifier)
                For ij As Integer = 0 To TBL.Rows.Count - 1
                    rowDummy = TBL.Rows(ij)
                    If rowDummy.delivery_date = rowSOB.promise_to_deliver Then
                        rowAcuan = TBL.Rows(ij)

                        'dilewat sementara, karena memungkinkan untuk terbalik, priority levelnya, misalnya order yang kedua high priority
                        'If ij < TBL.Rows.Count - 1 Then
                        '    rowDummy1 = TBL.Rows(ij + 1)

                        '    If Me.ObjPriorityLevel.GetUrutan(rowDummy1.delivery_priority) < _
                        '    Me.ObjPriorityLevel.GetUrutan(Me.ObjPriorityLevel.GetPrimary_ByDesc(Me.dgvData.Item(Me.ColPriorityLevel.Index, e.RowIndex).Value)) Then
                        '        MsgBox("Minimum priority is " + Me.ObjPriorityLevel.GetDesc(rowDummy1.delivery_priority), MsgBoxStyle.Information, "Information")
                        '        Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        '        Exit Sub
                        '    End If
                        'End If

                        'udate data primary priority level
                        Me.dgvData.Item(Me.ColPrimaryPriorityLevel.Index, e.RowIndex).Value = Me.ObjPriorityLevel.GetPrimary_ByDesc(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value)
                        'update ke dalam tabel delivery item
                        rowSOB.primary_priority_level = Me.dgvData.Item(Me.ColPrimaryPriorityLevel.Index, e.RowIndex).Value
                        rowSOB.urutan_priority_level = Me.ObjPriorityLevel.GetUrutan(Me.dgvData.Item(Me.ColPrimaryPriorityLevel.Index, e.RowIndex).Value)
                        'update ke dalam database
                        '1. jika ada item delivery di dalam db
                        'set primary + update ke tbl item delivery + update ke database + timestamp so
                        DTA.UpdateQuery_PrimaryDeliveryPriority(rowSOB.primary_priority_level, rowAcuan.identifier, rowAcuan.delivery_qty, rowAcuan.delivery_date)
                        DTA1.UpdateQuery_SetTimestamp(Now, rowAcuan.primary_transaksi)
                        '2. jika tidak ada item delivery dalam db -> bukan delivery data control -> tidak bisa diedit dari sini

                        'hitung ulang + posisikan ulang
                        Me.BackgroundWorker1.RunWorkerAsync()
                        Me.DataChanged = True
                    End If
                Next
            Else
                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA1.Dispose()
        End Try
    End Sub

    Private Sub EndEditPriorityLevel_PurchaseReq(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
        Dim TBL As New DseTransaksiShippingForecast.tbltrans_purchase_req_itemdeliveryDataTable
        Dim Identifier As String = ""
        Dim rowSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim rowAcuan, rowDummy As DseTransaksiShippingForecast.tbltrans_purchase_req_itemdeliveryRow
        Dim DTA1 As New DseTransaksiShippingForecastTableAdapters.tbltrans_purch_req_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            'tidak boleh string kosong
            If Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If

            'kalo sama dengan buffer tidak di proses
            If Me.Buffer.ToLower = Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToLower Then
                Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If

            'ambil data identifier
            Me.vTBLItemDelivery.DefaultView.RowFilter = "primary = " + Me.dgvData.Item(Me.ColPrimary.Index, e.RowIndex).Value.ToString
            If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                rowSOB = Me.vTBLItemDelivery.DefaultView(0).Row
                Identifier = rowSOB.identifier

                'pastikan priority level selalu lebih besar dari bawahnya
                'hanya cek jika bukan baris terakhir
                DTA.FillBy_ASC(TBL, Identifier)
                For ij As Integer = 0 To TBL.Rows.Count - 1
                    rowDummy = TBL.Rows(ij)
                    If rowDummy.delivery_date = rowSOB.promise_to_deliver Then
                        rowAcuan = TBL.Rows(ij)

                        'udate data primary priority level
                        Me.dgvData.Item(Me.ColPrimaryPriorityLevel.Index, e.RowIndex).Value = _
                        Me.ObjPriorityLevel.GetPrimary_ByDesc(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value)
                        'update ke dalam tabel delivery item
                        rowSOB.primary_priority_level = Me.dgvData.Item(Me.ColPrimaryPriorityLevel.Index, e.RowIndex).Value
                        rowSOB.urutan_priority_level = Me.ObjPriorityLevel.GetUrutan(Me.dgvData.Item(Me.ColPrimaryPriorityLevel.Index, e.RowIndex).Value)
                        'update ke dalam database
                        '1. jika ada item delivery di dalam db
                        'set primary + update ke tbl item delivery + update ke database + timestamp so
                        DTA.UpdateQuery_PrimaryDeliveryPriority(rowSOB.primary_priority_level, rowAcuan.identifier, rowAcuan.delivery_qty, rowAcuan.delivery_date)
                        DTA1.UpdateQuery_UpdateTimestamp(Now, rowAcuan.primary_transaksi)
                        '2. jika tidak ada item delivery dalam db -> bukan delivery data control -> tidak bisa diedit dari sini

                        'hitung ulang + posisikan ulang
                        Me.BackgroundWorker1.RunWorkerAsync()
                        Me.DataChanged = True
                    End If
                Next
            Else
                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA1.Dispose()
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Threading.Thread.Sleep(100)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Me.HitungUlang()
    End Sub

    Private Sub dgvData_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvData.CellContentClick
        Try
            If e.ColumnIndex <> -1 And e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.ColEdit.Index Then
                    If Me.dgvData.Item(Me.ColTipeItem.Index, e.RowIndex).Value = CInt(FrmTrans_MaterialPlanningDashboard_SF_DetailDelivery.EnumTipeItem.SOB) Then
                        If Me.dgvData.Item(Me.ColTipeInt.Index, e.RowIndex).Value = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB Then
                            If Me.ApakahDataMasihSamaDgnDB_SOB_EditLink(Me.dgvData.Item(Me.ColPrimary.Index, e.RowIndex).Value) = True Then
                                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_SF_Override_SalesETD") = True Then
                                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_SF_DOAll") = False Then
                                        If Me.dgvData.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = Me.UserCode Then
                                            Me.EditLink_ItemSOB(sender, e)
                                        Else
                                            MsgBox("This data is created by: " + Me.dgvData.Item(Me.ColCreatedBy.Index, e.RowIndex).Value.ToString + ", you cannot edit this data.", _
                                            MsgBoxStyle.Information, "Information")
                                        End If
                                    Else
                                        Me.EditLink_ItemSOB(sender, e)
                                    End If
                                End If
                            End If
                        Else
                            If Me.ApakahDataMasihSamaDgnDB_PurchReq_EditLink(Me.dgvData.Item(Me.ColPrimary.Index, e.RowIndex).Value) = True Then
                                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_SF_Override_SalesETD") = True Then
                                    'If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_SF_DOAll") = False Then
                                    '    If Me.dgvData.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = Me.UserCode Then
                                    '        Me.EditLink_ItemSOB(sender, e)
                                    '    Else
                                    '        MsgBox("This data is created by: " + Me.dgvData.Item(Me.ColCreatedBy.Index, e.RowIndex).Value.ToString + ", you cannot edit this data.", _
                                    '        MsgBoxStyle.Information, "Information")
                                    '    End If
                                    'Else
                                    Me.EditLink_ItemPurchReq(sender, e)
                                    'End If
                                End If
                            End If
                        End If
                    ElseIf Me.dgvData.Item(Me.ColTipeItem.Index, e.RowIndex).Value = CInt(FrmTrans_MaterialPlanningDashboard_SF_DetailDelivery.EnumTipeItem.POB) Then
                            If Me.ApakahDataMasihSamaDgDB_POB_EditLink(Me.dgvData.Item(Me.ColPrimary.Index, e.RowIndex).Value) = True Then
                                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_SF_Override_PurchaseETD") = True Then
                                    Me.EditLink_ItemPOB(sender, e)
                                Else
                                    MsgBox("Insufficient access, you cannot edit this data.", _
                                    MsgBoxStyle.Information, "Information")
                                End If
                            End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryMax_ItemPOB() As Double
        Dim rowpob As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow
        Dim Retval As Double = 0

        Try
            For Each rowpob In Me.vTBLItemPOB.Rows
                If Retval < rowpob.primary Then
                    Retval = rowpob.primary
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub EditLink_ItemPOB(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim ij As New FrmTrans_MaterialPlanningDashboard_SF_SetDeliveryPO
        Dim rowPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow
        Dim dummyIdentifier As String = ""
        Dim DTA1 As New DseTransaksiShippingForecastTableAdapters.tbltrans_po_itemdeliveryTableAdapter
        Dim DTAPOItem As New DseTransaksiShippingForecastTableAdapters.tbltrans_po_itemTableAdapter
        Dim TBLPOItem As New DseTransaksiShippingForecast.tbltrans_po_itemDataTable
        Dim rowPOItem As DseTransaksiShippingForecast.tbltrans_po_itemRow
        Dim DTAItemDeliveryPO As New DseTransaksiShippingForecastTableAdapters.tbltrans_po_itemdeliveryTableAdapter
        Dim TBLItemDeliveryPO As New DseTransaksiShippingForecast.tbltrans_po_itemdeliveryDataTable
        Dim rowItemDeliveryPO As DseTransaksiShippingForecast.tbltrans_po_itemdeliveryRow
        Dim rowItemPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow
        Dim DummyPrimaryTrans As Double = -1

        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTAPOItem.Connection.ConnectionString = DTA1.Connection.ConnectionString
            DTAItemDeliveryPO.Connection.ConnectionString = DTA1.Connection.ConnectionString

            Me.vTBLItemPOB.DefaultView.RowFilter = "primary = '" + Me.dgvData.Item(Me.ColPrimary.Index, e.RowIndex).Value.ToString + "'"
            If Me.vTBLItemPOB.DefaultView.Count = 0 Then
                MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                Exit Sub
            Else
                rowPOB = Me.vTBLItemPOB.DefaultView(0).Row
                dummyIdentifier = rowPOB.identifier
                DummyPrimaryTrans = rowPOB.primary_transaksi
                ij.VIdentifier = dummyIdentifier
                ij.vPrimaryCodeNo = rowPOB.primary_code_no
                ij.vRowPOBRef = rowPOB

                TBLPOItem.Rows.Clear()
                DTAPOItem.FillBy_Identifier(TBLPOItem, dummyIdentifier)
                rowPOItem = TBLPOItem.Rows(0)
                ij.vPrimaryUnitMulti = rowPOItem.Primary_Unit_Multi
                ij.vPrimaryUnitDefault = rowPOItem.Primary_Units
                ij.vRatioReference = rowPOItem.Ratio_Reference
                ij.vRatioDefault = rowPOItem.Ratio_Default
                ij.DefPrecMulti = Me.ObjProd.GetAngkaDecimal(rowPOItem.Primary_Product, rowPOItem.Primary_Unit_Multi)
                ij.DefPrecDefault = Me.ObjProd.GetAngkaDecimal(rowPOItem.Primary_Product)
                TBLPOItem.Rows.Clear()

                If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.DataChanged = True
                    'simpan perubahan ke dalam db: hapus + insert + set delivery date control pd item po
                    DTA1.DeleteQuery(dummyIdentifier)
                    For Each rowDGV As DataGridViewRow In ij.dgvDetail.Rows
                        If rowDGV.Index <> ij.dgvDetail.Rows.Count - 1 Then
                            DTA1.InsertQuery(rowPOB.primary_po, dummyIdentifier, rowDGV.Cells(ij.ColQty.Index).Value, _
                            rowDGV.Cells(ij.ColDeliveryDate.Index).Value, rowDGV.Cells(ij.ColRemarks.Index).Value, _
                            -1, rowDGV.Cells(ij.ColDelay.Index).Value, rowDGV.Cells(ij.ColQtyMulti.Index).Value)
                        End If
                    Next
                    If ij.dgvDetail.Rows.Count > 2 Then
                        Me.ObjPOItem.SetDeliveryDate2(dummyIdentifier, #1/1/1950#)
                    Else
                        Me.ObjPOItem.SetDeliveryDate2(dummyIdentifier, ij.dgvDetail.Item(ij.ColDeliveryDate.Index, 0).Value)
                    End If

                    'hapuskan semua item delivery yang identifiernya di edit
                    Me.vTBLItemPOB.DefaultView.RowFilter = "identifier = '" + dummyIdentifier + "' AND tipe = " + CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.POB).ToString
                    While Me.vTBLItemPOB.DefaultView.Count > 0
                        Me.vTBLItemPOB.Rows.Remove(Me.vTBLItemPOB.DefaultView(0).Row)
                    End While

                    'masukkan data item delivery baru
                    DTAPOItem.FillBy_Identifier(TBLPOItem, dummyIdentifier)
                    'pindahkan ke dalam data tabel yang akan digunakan untuk pengecekan otorisasi

                    'pindahkan ke dalam item pob
                    'cari primary po max dan min
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    If TBLPOItem.Rows.Count > 0 Then
                        'load item delivery
                        DTAItemDeliveryPO.FillBy_identifier(TBLItemDeliveryPO, dummyIdentifier)
                        TBLItemDeliveryPO.DefaultView.Sort = "[delivery date] DESC"
                        'isikan ke dalam tabel item pob
                        For Each rowPOItem In TBLPOItem.Rows
                            TBLItemDeliveryPO.DefaultView.RowFilter = "identifier = '" + rowPOItem.identifier + "'"
                            If TBLItemDeliveryPO.DefaultView.Count = 0 Then
                                'tidak ada delivery control -> set tanggal ke 1 jan 2200
                                rowItemPOB = Me.vTBLItemPOB.Newtbltrans_shipping_forecast_item_pobRow
                                With rowItemPOB
                                    .delivery_delay = 0
                                    .packing_list_date = #1/1/1950#
                                    .packing_list_no_rec = ""
                                    .po_date_rec = rowPOItem.PO_Date
                                    .po_number_rec = rowPOItem.PO_Number
                                    .primary = Me.GetPrimaryMax_ItemPOB + 1
                                    .primary_code_no = rowPOItem.Primary_Product
                                    .primary_packing_list = -1
                                    .primary_po = rowPOItem.Primary_Transaksi
                                    .primary_supplier = rowPOItem.Primary_Supplier
                                    .primary_supplier_rec = rowPOItem.Primary_Supplier
                                    .primary_transaksi = DummyPrimaryTrans
                                    .promise_to_deliver = #1/1/2200#
                                    .qty = rowPOItem.POB
                                    .remarks = ""
                                    .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.POB)
                                    .identifier = rowPOItem.identifier
                                End With
                                Me.vTBLItemPOB.Rows.Add(rowItemPOB)
                            Else
                                'ambil item delivery control mulai dari yang paling besar promise to deliverynya
                                While rowPOItem.POB > 0
                                    rowItemDeliveryPO = TBLItemDeliveryPO.DefaultView(0).Row
                                    rowItemPOB = Me.vTBLItemPOB.Newtbltrans_shipping_forecast_item_pobRow
                                    With rowItemPOB
                                        .delivery_delay = rowItemDeliveryPO.delay
                                        .packing_list_date = #1/1/1950#
                                        .packing_list_no_rec = ""
                                        .po_date_rec = rowPOItem.PO_Date
                                        .po_number_rec = rowPOItem.PO_Number
                                        .primary = Me.GetPrimaryMax_ItemPOB + 1
                                        .primary_code_no = rowPOItem.Primary_Product
                                        .primary_packing_list = -1
                                        .primary_po = rowPOItem.Primary_Transaksi
                                        .primary_supplier = rowPOItem.Primary_Supplier
                                        .primary_supplier_rec = rowPOItem.Primary_Supplier
                                        .primary_transaksi = DummyPrimaryTrans
                                        .promise_to_deliver = rowItemDeliveryPO.delivery_date
                                        If rowPOItem.POB > rowItemDeliveryPO.delivery_qty Then
                                            .qty = rowItemDeliveryPO.delivery_qty
                                            rowPOItem.POB -= rowItemDeliveryPO.delivery_qty
                                            rowItemDeliveryPO.delivery_qty = 0
                                        Else
                                            .qty = rowPOItem.POB
                                            rowPOItem.POB = 0
                                        End If
                                        .remarks = ""
                                        .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.POB)
                                        .identifier = rowPOItem.identifier
                                    End With
                                    Me.vTBLItemPOB.Rows.Add(rowItemPOB)

                                    'persiapan untuk looping selanjutnya
                                    If rowItemDeliveryPO.delivery_qty = 0 Then
                                        TBLItemDeliveryPO.Rows.Remove(rowItemDeliveryPO)
                                    End If
                                End While
                            End If
                        Next
                    End If

                    'hitung ulang
                    Me.HitungUlang()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            DTA1.Dispose()
            DTAPOItem.Dispose()
            TBLPOItem.Dispose()
            DTAItemDeliveryPO.Dispose()
            TBLItemDeliveryPO.Dispose()
        End Try
    End Sub

    Private Sub EditLink_ItemSOB(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim FrmSOB As New FrmTrans_MaterialPlanningDashboard_SF_SetDeliverySO
        Dim rowSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_so_itemdelivery1TableAdapter
        Dim DTASOItem As New DseTransaksiShippingForecastTableAdapters.tbltrans_so_itemTableAdapter
        Dim TBLSOItem As New DseTransaksiShippingForecast.tbltrans_so_itemDataTable
        Dim rowSOItem As DseTransaksiShippingForecast.tbltrans_so_itemRow
        Dim DummyIdentifier As String = ""
        Dim DTASOItemDelivery As New DseTransaksiShippingForecastTableAdapters.tbltrans_so_itemdeliveryTableAdapter
        Dim TBLSOItemDelivery As New DseTransaksiShippingForecast.tbltrans_so_itemdeliveryDataTable
        Dim rowSOItemDelivery As DseTransaksiShippingForecast.tbltrans_so_itemdeliveryRow
        Dim rowItemSO_SF As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim PrimaryPriorityLevelDefault As Double = Me.ObjPriorityLevel.GetPrimaryDefault
        Dim UrutanPriodityLevelDefault As Integer = Me.ObjPriorityLevel.GetUrutan(PrimaryPriorityLevelDefault)
        Dim PrimaryTransaksi As Double = -1
        Dim DTABooking As New DseTransaksiShippingForecastTableAdapters.tbltrans_booking_invitemTableAdapter
        Dim TBLBooking As New DseTransaksiShippingForecast.tbltrans_booking_invitemDataTable
        Dim rowBooking As DseTransaksiShippingForecast.tbltrans_booking_invitemRow
        Dim IndexAcuan As Double = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTASOItem.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTASOItemDelivery.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTABooking.Connection.ConnectionString = Core_Module.GetConstring

            FrmSOB.vPrimaryCodeNo = Me.vPrimaryCodeNoTarget

            Me.vTBLItemDelivery.DefaultView.RowFilter = "primary = " + Me.dgvData.Item(Me.ColPrimary.Index, e.RowIndex).Value.ToString
            If Me.vTBLItemDelivery.DefaultView.Count = 0 Then
                MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                Exit Sub
            Else
                rowSOB = Me.vTBLItemDelivery.DefaultView(0).Row
                FrmSOB.VIdentifier = rowSOB.identifier
                DummyIdentifier = rowSOB.identifier
                PrimaryTransaksi = rowSOB.primary_transaksi
                FrmSOB.vPrimaryCodeNo = rowSOB.primary_code_no
                FrmSOB.vReadOnly = False
                FrmSOB.vRowSOBRef = rowSOB

                'isikan data multi unit
                DTASOItem.FillBy_Identifier(TBLSOItem, DummyIdentifier)
                rowSOItem = TBLSOItem.Rows(0)
                FrmSOB.vPrimaryUnitMulti = rowSOItem.Primary_Units_Multi
                FrmSOB.vPrimaryUnitDefault = rowSOItem.Primary_Units
                FrmSOB.vRatioReference = rowSOItem.Ratio_Reference
                FrmSOB.vRatioDefault = rowSOItem.Ratio_Default
                FrmSOB.vDefPrecMulti = Me.ObjProd.GetAngkaDecimal(rowSOItem.Primary_Product, rowSOItem.Primary_Units_Multi)
                FrmSOB.vDefPrecDefault = Me.ObjProd.GetAngkaDecimal(rowSOItem.Primary_Product)
                TBLSOItem.Rows.Clear()

                If FrmSOB.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.DataChanged = True
                    'simpan perubahan ke dalam db: hapus + insert + set delivery date control pd item so
                    DTA.DeleteQuery(rowSOB.identifier)
                    For Each rowDgv As DataGridViewRow In FrmSOB.dgvDetail.Rows
                        If rowDgv.Index <> FrmSOB.dgvDetail.Rows.Count - 1 Then
                            DTA.InsertQuery(rowSOB.primary_so, rowSOB.identifier, rowDgv.Cells(FrmSOB.ColQty.Index).Value, rowDgv.Cells(FrmSOB.ColDeliveryDate.Index).Value, _
                            rowDgv.Cells(FrmSOB.ColRemarks.Index).Value, rowDgv.Cells(FrmSOB.ColPrimaryPriority.Index).Value, rowDgv.Cells(FrmSOB.ColQtyMulti.Index).Value)
                        End If
                    Next
                    If FrmSOB.dgvDetail.Rows.Count > 2 Then
                        Me.ObjSOItem.SetDeliveryDateControl(#1/1/1950#, rowSOB.identifier)
                    Else
                        Me.ObjSOItem.SetDeliveryDateControl(FrmSOB.dgvDetail.Item(FrmSOB.ColDeliveryDate.Index, 0).Value, rowSOB.identifier)
                    End If
                    'hapuskan semua item delivery yang identifiernya di edit
                    Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + rowSOB.identifier + "'"
                    While Me.vTBLItemDelivery.DefaultView.Count > 0
                        Me.vTBLItemDelivery.Rows.Remove(Me.vTBLItemDelivery.DefaultView(0).Row)
                    End While

                    'masukkan data item delivery baru
                    DTASOItem.FillBy_Identifier(TBLSOItem, DummyIdentifier)
                    If TBLSOItem.Rows.Count > 0 Then
                        DTASOItemDelivery.FillBy_Identifier(TBLSOItemDelivery, DummyIdentifier)
                        TBLSOItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                        For Each rowSOItem In TBLSOItem.Rows
                            TBLSOItemDelivery.DefaultView.RowFilter = "identifier = '" + rowSOItem.identifier + "'"

                            If TBLSOItemDelivery.DefaultView.Count = 0 Then
                                'tambahkan ke dalam item so sesuai sob
                                rowItemSO_SF = Me.vTBLItemDelivery.Newtbltrans_shipping_forecast_itemdeliveryRow
                                With rowItemSO_SF
                                    .create_do = 0
                                    'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery
                                    'dan 0 untuk qty available
                                    .qty_available = 0
                                    .estimate_delivery = #1/1/2200#
                                    .primary = Me.GetPrimaryMax_ItemSOB + 1
                                    .primary_code_no = rowSOItem.Primary_Product
                                    .primary_cust = rowSOItem.Primary_Customer
                                    .primary_cust_rec = rowSOItem.Primary_Customer
                                    .primary_priority_level = PrimaryPriorityLevelDefault
                                    .primary_so = rowSOItem.Primary_Transaksi
                                    .primary_transaksi = PrimaryTransaksi
                                    .primary_unit = rowSOItem.Primary_Units
                                    'tidak ada delivery control -> set tanggal ke 1 jan 2200
                                    .promise_to_deliver = #1/1/2200#
                                    .qty_sob = rowSOItem.SOB
                                    .remark = ""
                                    .so_created_by = rowSOItem.Created_By
                                    .so_date_rec = rowSOItem.Tanggal_Transaksi
                                    .so_number_rec = rowSOItem.No_Transaksi
                                    .urutan_priority_level = UrutanPriodityLevelDefault
                                    .customer_name = rowSOItem.Customer_Name
                                    .Product_Code = rowSOItem.Product_Code
                                    .Description = rowSOItem.Description
                                    .so_number = rowSOItem.No_Transaksi
                                    .so_date = rowSOItem.Tanggal_Transaksi
                                    .Cust_PO_No = rowSOItem.Cust_PO_No
                                    ._Decimal = rowSOItem._Decimal
                                    .Unit_Code = rowSOItem.Unit_Code
                                    .Primary_Category = rowSOItem.Primary_Category
                                    .Primary_Sub_Brand = rowSOItem.Primary_Sub_Brand
                                    .Primary_Brand = rowSOItem.Primary_Brand
                                    .primary_product_group = rowSOItem.Primary_Product_Group
                                    .identifier = rowSOItem.identifier
                                    .booked = 0
                                    .qty_ready_stock = 0
                                    .no_booking = ""
                                    .qty_booked = 0
                                    .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB)
                                    .tipe_product = rowSOItem.Type_Product
                                End With
                                Me.vTBLItemDelivery.Rows.Add(rowItemSO_SF)
                            Else
                                'ambil item delivery control mulai dari yang paling besar promise to deliverynya
                                While rowSOItem.SOB > 0
                                    rowSOItemDelivery = TBLSOItemDelivery.DefaultView(0).Row
                                    rowItemSO_SF = Me.vTBLItemDelivery.Newtbltrans_shipping_forecast_itemdeliveryRow
                                    With rowItemSO_SF
                                        .create_do = 0
                                        'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery dan 0 untuk qty availabel
                                        .estimate_delivery = #1/1/2200#
                                        .qty_available = 0
                                        .primary = Me.GetPrimaryMax_ItemSOB + 1
                                        .primary_code_no = rowSOItem.Primary_Product
                                        .primary_cust = rowSOItem.Primary_Customer
                                        .primary_cust_rec = rowSOItem.Primary_Customer
                                        .primary_priority_level = rowSOItemDelivery.delivery_priority
                                        .urutan_priority_level = rowSOItemDelivery.urutan
                                        .primary_so = rowSOItem.Primary_Transaksi
                                        .primary_transaksi = PrimaryTransaksi
                                        .primary_unit = rowSOItem.Primary_Units
                                        'tidak ada delivery control -> set tanggal ke 1 jan 2200
                                        .promise_to_deliver = rowSOItemDelivery.delivery_date
                                        If rowSOItem.SOB > rowSOItemDelivery.delivery_qty Then
                                            .qty_sob = rowSOItemDelivery.delivery_qty
                                            rowSOItem.SOB -= rowSOItemDelivery.delivery_qty
                                            rowSOItemDelivery.delivery_qty = 0
                                            rowSOItemDelivery.delivery_qty_multi = 0
                                        Else
                                            .qty_sob = rowSOItem.SOB
                                            rowSOItem.SOB = 0
                                        End If
                                        .remark = ""
                                        .so_created_by = rowSOItem.Created_By
                                        .so_date_rec = rowSOItem.Tanggal_Transaksi
                                        .so_number_rec = rowSOItem.No_Transaksi
                                        .customer_name = rowSOItem.Customer_Name
                                        .Product_Code = rowSOItem.Product_Code
                                        .Description = rowSOItem.Description
                                        .so_number = rowSOItem.No_Transaksi
                                        .so_date = rowSOItem.Tanggal_Transaksi
                                        .Cust_PO_No = rowSOItem.Cust_PO_No
                                        ._Decimal = rowSOItem._Decimal
                                        .Unit_Code = rowSOItem.Unit_Code
                                        .Primary_Category = rowSOItem.Primary_Category
                                        .Primary_Sub_Brand = rowSOItem.Primary_Sub_Brand
                                        .Primary_Brand = rowSOItem.Primary_Brand
                                        .primary_product_group = rowSOItem.Primary_Product_Group
                                        .identifier = rowSOItem.identifier
                                        .booked = 0
                                        .no_booking = ""
                                        .qty_booked = 0
                                        .qty_ready_stock = 0
                                        .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB)
                                        .tipe_product = rowSOItem.Type_Product
                                        If rowSOItemDelivery.delivery_qty = 0 Then
                                            TBLSOItemDelivery.Rows.Remove(rowSOItemDelivery)
                                        End If
                                    End With
                                    Me.vTBLItemDelivery.Rows.Add(rowItemSO_SF)
                                End While
                            End If
                        Next
                    End If

                    'tandain item sob yang ada inventory bookingnya
                    DTABooking.FillBy_PrimaryCodeNo(TBLBooking, Me.vPrimaryCodeNoTarget)
                    For Each rowBooking In TBLBooking.Rows
                        Me.vTBLItemDelivery.DefaultView.RowFilter = "[primary code no] = " + rowBooking.Primary_Product.ToString + " AND [primary so] = " + rowBooking.Primary_SO.ToString
                        Me.vTBLItemDelivery.DefaultView.Sort = "[promise to deliver] ASC, [urutan priority level] ASC, [so date rec] ASC, [primary so] ASC"
                        IndexAcuan = 0
                        While IndexAcuan < Me.vTBLItemDelivery.DefaultView.Count And rowBooking.Qty > 0
                            rowItemSO_SF = Me.vTBLItemDelivery.DefaultView(IndexAcuan).Row

                            If rowItemSO_SF.qty_sob - rowItemSO_SF.qty_booked >= rowBooking.Qty Then
                                rowItemSO_SF.qty_booked += rowBooking.Qty
                                rowBooking.Qty = 0
                            Else
                                rowBooking.Qty -= (rowItemSO_SF.qty_sob - rowItemSO_SF.qty_booked)
                                rowItemSO_SF.qty_booked = rowItemSO_SF.qty_sob
                                IndexAcuan += 1
                            End If

                            rowItemSO_SF.booked = 1
                            If rowItemSO_SF.no_booking = "" Then
                                rowItemSO_SF.no_booking = rowBooking.No_Transaksi
                            Else
                                If rowItemSO_SF.no_booking.Contains(rowBooking.No_Transaksi) = False Then
                                    rowItemSO_SF.no_booking = rowItemSO_SF.no_booking + ", " + rowBooking.No_Transaksi
                                End If
                            End If
                        End While
                    Next

                    'hitung ulang
                    Me.HitungUlang()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTABooking.Dispose()
            TBLBooking.Dispose()
            FrmSOB.Dispose()
            DTA.Dispose()
            DTASOItem.Dispose()
            TBLSOItem.Dispose()
            DTASOItemDelivery.Dispose()
            TBLSOItemDelivery.Dispose()
        End Try
    End Sub

    Private Sub EditLink_ItemPurchReq(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim FrmPurcReqB As New FrmTrans_MaterialPlanningDashboard_SF_SetDeliveryPurcReq
        Dim vRowPurcReqBRef As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow

        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
        Dim DTAPurcReqItem As New DseTransaksiShippingForecastTableAdapters.tbltrans_purch_req_itemTableAdapter
        Dim TBLPurcReqItem As New DseTransaksiShippingForecast.tbltrans_purch_req_itemDataTable
        Dim rowPurcReqItem As DseTransaksiShippingForecast.tbltrans_purch_req_itemRow
        Dim DummyIdentifier As String = ""
        Dim DTAPurcReqItemDelivery As New DseTransaksiShippingForecastTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
        Dim TBLPurcReqItemDelivery As New DseTransaksiShippingForecast.tbltrans_purchase_req_itemdeliveryDataTable
        Dim rowPurcReqItemDelivery As DseTransaksiShippingForecast.tbltrans_purchase_req_itemdeliveryRow
        Dim rowItemPurcReq_SF As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow

        Dim PrimaryPriorityLevelDefault As Double = Me.ObjPriorityLevel.GetPrimaryDefault
        Dim UrutanPriodityLevelDefault As Integer = Me.ObjPriorityLevel.GetUrutan(PrimaryPriorityLevelDefault)
        Dim PrimaryTransaksi As Double = -1
        Dim IndexAcuan As Double = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAPurcReqItem.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAPurcReqItemDelivery.Connection.ConnectionString = DTA.Connection.ConnectionString

            FrmPurcReqB.vPrimaryCodeNo = Me.vPrimaryCodeNoTarget
            Me.vTBLItemDelivery.DefaultView.RowFilter = "primary = " + Me.dgvData.Item(Me.ColPrimary.Index, e.RowIndex).Value.ToString
            If Me.vTBLItemDelivery.DefaultView.Count = 0 Then
                MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                Exit Sub
            Else
                vRowPurcReqBRef = Me.vTBLItemDelivery.DefaultView(0).Row
                FrmPurcReqB.VIdentifier = vRowPurcReqBRef.identifier
                DummyIdentifier = vRowPurcReqBRef.identifier
                PrimaryTransaksi = vRowPurcReqBRef.primary_transaksi
                FrmPurcReqB.vPrimaryCodeNo = vRowPurcReqBRef.primary_code_no
                FrmPurcReqB.vReadOnly = False
                FrmPurcReqB.vRowPurcReqBRef = vRowPurcReqBRef

                If FrmPurcReqB.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.DataChanged = True
                    'simpan perubahan ke dalam db: hapus + insert + set delivery date control pd item PurcReq
                    DTA.DeleteQuery(vRowPurcReqBRef.identifier)
                    For Each rowDgv As DataGridViewRow In FrmPurcReqB.dgvDetail.Rows
                        If rowDgv.Index <> FrmPurcReqB.dgvDetail.Rows.Count - 1 Then
                            DTA.InsertQuery(vRowPurcReqBRef.primary_so, vRowPurcReqBRef.identifier, rowDgv.Cells(FrmPurcReqB.ColQty.Index).Value, _
                            rowDgv.Cells(FrmPurcReqB.ColDeliveryDate.Index).Value, _
                            rowDgv.Cells(FrmPurcReqB.ColRemarks.Index).Value, rowDgv.Cells(FrmPurcReqB.ColPrimaryPriority.Index).Value)
                        End If
                    Next
                    If FrmPurcReqB.dgvDetail.Rows.Count > 2 Then
                        Me.ObjPurcReq.SetDeliveryDateControl(#1/1/1950#, vRowPurcReqBRef.identifier)
                    Else
                        Me.ObjPurcReq.SetDeliveryDateControl(FrmPurcReqB.dgvDetail.Item(FrmPurcReqB.ColDeliveryDate.Index, 0).Value, vRowPurcReqBRef.identifier)
                    End If

                    'hapuskan semua item delivery yang identifiernya di edit
                    Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + vRowPurcReqBRef.identifier + "'"
                    While Me.vTBLItemDelivery.DefaultView.Count > 0
                        Me.vTBLItemDelivery.Rows.Remove(Me.vTBLItemDelivery.DefaultView(0).Row)
                    End While

                    'masukkan data item delivery baru
                    DTAPurcReqItem.FillBy_Identifier(TBLPurcReqItem, DummyIdentifier)
                    If TBLPurcReqItem.Rows.Count > 0 Then
                        DTAPurcReqItemDelivery.FillBy_Identifier(TBLPurcReqItemDelivery, DummyIdentifier)
                        TBLPurcReqItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                        For Each rowPurcReqItem In TBLPurcReqItem.Rows
                            TBLPurcReqItemDelivery.DefaultView.RowFilter = "identifier = '" + rowPurcReqItem.identifier + "'"

                            If TBLPurcReqItemDelivery.DefaultView.Count = 0 Then
                                'tambahkan ke dalam item PurcReq sesuai PurcReqb
                                rowItemPurcReq_SF = Me.vTBLItemDelivery.Newtbltrans_shipping_forecast_itemdeliveryRow
                                With rowItemPurcReq_SF
                                    .create_do = 0
                                    'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery
                                    'dan 0 untuk qty available
                                    .qty_available = 0
                                    .estimate_delivery = #1/1/2200#
                                    .primary = Me.GetPrimaryMax_ItemSOB + 1
                                    .primary_code_no = rowPurcReqItem.primary_code_no
                                    .primary_cust = rowPurcReqItem.primary_ref_party
                                    .primary_cust_rec = rowPurcReqItem.primary_ref_party
                                    .primary_priority_level = PrimaryPriorityLevelDefault
                                    .primary_so = rowPurcReqItem.primary_transaksi
                                    .primary_transaksi = PrimaryTransaksi
                                    .primary_unit = rowPurcReqItem.primary_unit
                                    'tidak ada delivery control -> set tanggal ke 1 jan 2200
                                    .promise_to_deliver = #1/1/2200#
                                    .qty_sob = rowPurcReqItem.OB
                                    .remark = rowPurcReqItem.remarks_item
                                    .so_created_by = rowPurcReqItem.created_by
                                    .so_date_rec = rowPurcReqItem.tanggal
                                    .so_number_rec = rowPurcReqItem.no_transaksi
                                    .urutan_priority_level = UrutanPriodityLevelDefault
                                    .customer_name = rowPurcReqItem.Ref_Party_Desc
                                    .Product_Code = rowPurcReqItem.Product_Code
                                    .Description = rowPurcReqItem.Description
                                    .so_number = rowPurcReqItem.no_transaksi
                                    .so_date = rowPurcReqItem.tanggal
                                    .Cust_PO_No = ""
                                    ._Decimal = rowPurcReqItem._Decimal
                                    .Unit_Code = rowPurcReqItem.Unit_Code
                                    .Primary_Category = rowPurcReqItem.Primary_Category
                                    .Primary_Sub_Brand = rowPurcReqItem.Primary_Sub_Brand
                                    .Primary_Brand = rowPurcReqItem.Primary_Brand
                                    .primary_product_group = rowPurcReqItem.Primary_Product_Group
                                    .identifier = rowPurcReqItem.identifier
                                    .booked = 0
                                    .no_booking = ""
                                    .qty_booked = 0
                                    .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req)
                                    .tipe_product = rowPurcReqItem.Type
                                End With
                                Me.vTBLItemDelivery.Rows.Add(rowItemPurcReq_SF)
                            Else
                                'ambil item delivery control mulai dari yang paling besar promise to deliverynya
                                While rowPurcReqItem.OB > 0
                                    rowPurcReqItemDelivery = TBLPurcReqItemDelivery.DefaultView(0).Row
                                    rowItemPurcReq_SF = Me.vTBLItemDelivery.Newtbltrans_shipping_forecast_itemdeliveryRow
                                    With rowItemPurcReq_SF
                                        .create_do = 0
                                        'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery dan 0 untuk qty availabel
                                        .estimate_delivery = #1/1/2200#
                                        .qty_available = 0
                                        .qty_ready_stock = 0
                                        .primary = Me.GetPrimaryMax_ItemSOB + 1
                                        .primary_code_no = rowPurcReqItem.primary_code_no
                                        .primary_cust = rowPurcReqItem.primary_ref_party
                                        .primary_cust_rec = rowPurcReqItem.primary_ref_party
                                        .primary_priority_level = rowPurcReqItemDelivery.delivery_priority
                                        .urutan_priority_level = rowPurcReqItemDelivery.urutan
                                        .primary_so = rowPurcReqItem.primary_transaksi
                                        .primary_transaksi = PrimaryTransaksi
                                        .primary_unit = rowPurcReqItem.primary_unit
                                        'tidak ada delivery control -> set tanggal ke 1 jan 2200
                                        .promise_to_deliver = rowPurcReqItemDelivery.delivery_date
                                        If rowPurcReqItem.OB > rowPurcReqItemDelivery.delivery_qty Then
                                            .qty_sob = rowPurcReqItemDelivery.delivery_qty
                                            rowPurcReqItem.OB -= rowPurcReqItemDelivery.delivery_qty
                                            rowPurcReqItemDelivery.delivery_qty = 0
                                        Else
                                            .qty_sob = rowPurcReqItem.OB
                                            rowPurcReqItem.OB = 0
                                        End If
                                        .remark = rowPurcReqItem.remarks_item
                                        .so_created_by = rowPurcReqItem.created_by
                                        .so_date_rec = rowPurcReqItem.tanggal
                                        .so_number_rec = rowPurcReqItem.no_transaksi
                                        .customer_name = rowPurcReqItem.Ref_Party_Desc
                                        .Product_Code = rowPurcReqItem.Product_Code
                                        .Description = rowPurcReqItem.Description
                                        .so_number = rowPurcReqItem.no_transaksi
                                        .so_date = rowPurcReqItem.tanggal
                                        .Cust_PO_No = ""
                                        ._Decimal = rowPurcReqItem._Decimal
                                        .Unit_Code = rowPurcReqItem.Unit_Code
                                        .Primary_Category = rowPurcReqItem.Primary_Category
                                        .Primary_Sub_Brand = rowPurcReqItem.Primary_Sub_Brand
                                        .Primary_Brand = rowPurcReqItem.Primary_Brand
                                        .primary_product_group = rowPurcReqItem.Primary_Product_Group
                                        .identifier = rowPurcReqItem.identifier
                                        .booked = 0
                                        .no_booking = ""
                                        .qty_booked = 0
                                        .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req)
                                        .tipe_product = rowPurcReqItem.Type
                                        If rowPurcReqItemDelivery.delivery_qty = 0 Then
                                            TBLPurcReqItemDelivery.Rows.Remove(rowPurcReqItemDelivery)
                                        End If
                                    End With
                                    Me.vTBLItemDelivery.Rows.Add(rowItemPurcReq_SF)
                                End While
                            End If
                        Next
                    End If

                    'hitung ulang
                    Me.HitungUlang()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmPurcReqB.Dispose()
            DTA.Dispose()
            DTAPurcReqItem.Dispose()
            TBLPurcReqItem.Dispose()
            DTAPurcReqItemDelivery.Dispose()
            TBLPurcReqItemDelivery.Dispose()
        End Try
    End Sub

    Private Function GetPrimaryMax_ItemSOB() As Double
        Dim rowsob As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim Retval As Double = 0

        Try
            For Each rowsob In Me.vTBLItemDelivery.Rows
                If Retval < rowsob.primary Then
                    Retval = rowsob.primary
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Me.HitungUlang()
    End Sub
End Class