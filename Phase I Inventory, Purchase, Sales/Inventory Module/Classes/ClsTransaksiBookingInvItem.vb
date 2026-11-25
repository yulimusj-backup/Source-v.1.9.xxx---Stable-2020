Option Explicit On
Option Strict Off

Public Class ClsTransaksiBookingInvItem
    Public Sub HapusBookingByValidDate(ByVal vTanggalSampai As Date)
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery_ByValidDate(vTanggalSampai)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function IsiTBLItem(ByVal vPrimarySO As Double, _
    ByVal TBL As DseTransaksiBookingInv.tbltrans_booking_invitem1DataTable) As Boolean
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitem1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBL.Rows.Clear()
            DTA.Fill(TBL, vPrimarySO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' fungsi ini digunakan untuk cek stock. Mengembalikan total barang yang primarynya diberikan
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJml(ByVal vPrimaryCodeNo As Double) As Double
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vPrimaryCodeNo).HasValue = True Then
                Return DTA.ScalarQuery_GetJml(vPrimaryCodeNo).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' fungsi ini digunakan untuk cek stock. Mengembalikan total barang yang primarynya diberikan
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJmlByPrimaryProductNBooking(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryBooking As Double) As Double
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitem2TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryProductNBooking(vPrimaryCodeNo, vPrimaryBooking).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByPrimaryProductNBooking(vPrimaryCodeNo, vPrimaryBooking).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml2(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryBooking As Double, _
    ByVal vPrimarySO As Double) As Double
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitem2TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryProductNBooking(vPrimaryCodeNo, vPrimaryBooking).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByPrimaryProductNBooking(vPrimaryCodeNo, vPrimaryBooking).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlOtherBookingInv(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryBooking As Double, _
    ByVal vPrimarySO As Double) As Double
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitem2TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_getJmlOtherBookingInv(vPrimaryCodeNo, vPrimaryBooking, vPrimarySO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlOtherBookingInv(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryBooking As Double, _
    ByVal vPrimarySO As Double, ByVal vBatch As String, ByVal vPrimaryLocation As Double, ByVal vSecQty As Double) As Double
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitem2TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJmlOther2(vPrimaryCodeNo, vPrimaryBooking, vPrimarySO, _
            vBatch, vSecQty, vPrimaryLocation).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlOtherBookingInv2(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryBooking As Double, _
    ByVal vBatch As String, ByVal vPrimaryLocation As Double, ByVal vSecQty As Double) As Double
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitem2TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJmlOther3(vPrimaryCodeNo, vPrimaryBooking, _
            vBatch, vSecQty, vPrimaryLocation).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan jumlah booking yang tidak termasuk untuk SO yang diberikan primary-nya
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <param name="vPrimarySO"></param>
    ''' <param name="vBatch"></param>
    ''' <param name="vSecQty"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJmlOtherBookingInv3(ByVal vPrimaryCodeNo As Double, ByVal vPrimarySO As Double, _
    ByVal vBatch As String, ByVal vSecQty As Double) As Double
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitem2TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJmlOther4(vPrimaryCodeNo, _
            vBatch, vSecQty, vPrimarySO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDgv(ByVal vDGV As DataGridView, ByVal TBL As _
    DseTransaksiBookingInv.tbltrans_bookinginv_item_viewDataTable, Optional ByVal _
    vPrimaryProduct As Double = -1) As Boolean
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_bookinginv_item_viewTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vPrimaryProduct = -1 Then
                DTA.Fill(TBL)
            Else
                DTA.FillBy_PrimaryProduct(TBL, vPrimaryProduct)
            End If
            vDGV.DataSource = TBL.DefaultView
            vDGV.Columns(1).DefaultCellStyle.Format = "dd MMM yyyy"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDgv1(ByVal vDGV As DataGridView, ByVal TBL As _
DseTransaksiBookingInv.tbltrans_bookinginv_item_view1DataTable, Optional ByVal _
vPrimaryProduct As Double = -1) As Boolean
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_bookinginv_item_view1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vPrimaryProduct = -1 Then
                DTA.Fill(TBL)
            Else
                DTA.FillBy_PrimaryProduct(TBL, vPrimaryProduct)
            End If
            vDGV.DataSource = TBL.DefaultView
            vDGV.Columns(1).DefaultCellStyle.Format = "dd MMM yyyy"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsikanTBLLocation(ByVal TBL As DseTransaksiBookingInv.DataTable5DataTable, ByVal vPrimaryProduct As Double, _
    ByVal vPrimaryTransaksiBooking As Double, ByVal vBatchNo As String)
        Dim DTA As New DseTransaksiBookingInvTableAdapters.DataTable5TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy(TBL, vPrimaryTransaksiBooking, vPrimaryProduct, vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLProductBySO(ByVal vPrimaryBooking As Double, ByVal TBL As DseTransaksiBookingInv.DataTable22DataTable, _
    ByVal vPrimarySO As Double)
        Dim DTA As New DseTransaksiBookingInvTableAdapters.DataTable22TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimarySONBooking(TBL, vPrimarySO, vPrimaryBooking)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLBatchNo(ByVal vPrimaryBooking As Double, ByVal vPrimaryProduct As Double, _
    ByVal TBL As DseTransaksiBookingInv.tbltrans_booking_invitem3DataTable)
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitem3TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryProduct_PrimaryBooking(TBL, vPrimaryBooking, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLSecondaryQty(ByVal vPrimaryBooking As Double, ByVal vPrimaryProduct As Double, _
    ByVal vBatchNo As String, ByVal vPrimaryLocation As Double, _
    ByVal TBL As DseTransaksiBookingInv.tbltrans_booking_invitem4DataTable)
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitem4TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryBooking, vPrimaryProduct, vPrimaryLocation, vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetJml(ByVal vPrimaryBooking As Double, ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double) As Integer
        Dim DTA As New DseTransaksiBookingInvTableAdapters.DataTable22TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vPrimarySO, vPrimaryProduct, vPrimaryBooking).HasValue = True Then
                Return DTA.ScalarQuery_GetJml(vPrimarySO, vPrimaryProduct, vPrimaryBooking).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml3(ByVal vPrimaryBooking As Double, ByVal vPrimaryProduct As Double, _
    ByVal vSecondaryQty As Double, ByVal vBatchNo As String) As Double
        Dim DTA As New DseTransaksiBookingInvTableAdapters.DataTable22TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_getJmlByPrimaryBookingPrimaryProductBatchNoSecQty(vPrimaryBooking, vPrimaryProduct, vBatchNo, vSecondaryQty).HasValue = True Then
                Return DTA.ScalarQuery_getJmlByPrimaryBookingPrimaryProductBatchNoSecQty(vPrimaryBooking, vPrimaryProduct, vBatchNo, vSecondaryQty).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vPrimaryProduct As Double, ByVal vPrimarySO As Double, _
    ByVal vSecondaryQty As Double, ByVal vBatchNo As String, ByVal vPrimaryLocation As Double) As Double
        Dim DTA As New DseTransaksiBookingInvTableAdapters.DataTable22TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJml1(vPrimaryProduct, vBatchNo, _
            vPrimarySO, vSecondaryQty, vPrimaryLocation).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml4(ByVal vPrimaryProduct As Double, ByVal vPrimarySO As Double, _
    ByVal vSecondaryQty As Double, ByVal vBatchNo As String) As Double
        Dim DTA As New DseTransaksiBookingInvTableAdapters.DataTable22TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJml2(vPrimaryProduct, vBatchNo, _
            vPrimarySO, vSecondaryQty).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml5(ByVal vPrimaryProduct As Double, _
    ByVal vSecondaryQty As Double, ByVal vBatchNo As String) As Double
        Dim DTA As New DseTransaksiBookingInvTableAdapters.DataTable22TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQueryGetJml3(vPrimaryProduct, vBatchNo, _
            vSecondaryQty).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vPrimaryProduct As Double, ByVal vSecQty As Double) As Double
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml2(vPrimaryProduct, vSecQty).HasValue = True Then
                Return DTA.ScalarQuery_GetJml2(vPrimaryProduct, vSecQty).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGVCariDetail(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal i As FrmTransaksiBookingInv) As Boolean
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tblCariDetailItemTableAdapter
        Dim TBL As New DseTransaksiBookingInv.tblCariDetailItemDataTable
        Dim row As DseTransaksiBookingInv.tblCariDetailItemRow
        Dim Sec_Units As String = ""
        Dim SecQty As String = ""
        Dim ExpiryDate As String = ""
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowdgv As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.FillBy_PrimaryBooking(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    If row.Primary_Sec_Unit <> -1 Then
                        Sec_Units = row.Unit_Code_Sec_Qty
                    Else
                        Sec_Units = ""
                    End If
                    If row.Expiry_Date = #1/1/1950# Then
                        ExpiryDate = "-"
                    End If
                    If row.Sec_Qty = "-1" Then
                        SecQty = "-"
                    End If

                    With row
                        vDGVItem.Rows.Add()
                        rowdgv = i.dgvItem.Rows(vDGVItem.Rows.Count - 1)

                        rowdgv.Cells(i.ColNo.Index).Value = "0"
                        rowdgv.Cells(i.ColCodeNo.Index).Value = .Product_Code
                        rowdgv.Cells(i.ColDescription.Index).Value = .Description
                        rowdgv.Cells(i.ColQty.Index).Value = .Qty
                        rowdgv.Cells(i.ColUnit.Index).Value = .Unit_Code
                        rowdgv.Cells(i.ColSecondaryQty.Index).Value = .Sec_Qty
                        rowdgv.Cells(i.ColSecondaryUnit.Index).Value = .Unit_Code_Sec_Qty
                        rowdgv.Cells(i.ColBatch.Index).Value = .Batch_No
                        rowdgv.Cells(i.ColExpiredDate.Index).Value = .Expiry_Date
                        rowdgv.Cells(i.ColValidUntil.Index).Value = .Valid_Until
                        rowdgv.Cells(i.ColRemark.Index).Value = .Remarks
                        rowdgv.Cells(i.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rowdgv.Cells(i.ColPrimaryUnits.Index).Value = .Primary_Unit
                        rowdgv.Cells(i.ColPrimarySecondaryUnits.Index).Value = .Primary_Sec_Unit
                        rowdgv.Cells(i.ColPrimaryLocation.Index).Value = .Primary_Location
                        rowdgv.Cells(i.ColLocation.Index).Value = .Location
                        rowdgv.Cells(i.ColWarehouse.Index).Value = .Warehouse
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
        End Try
    End Function

    Public Function GetQtyByPrimarySO_n_Product(ByVal vPrimaryProduct As Double, _
    ByVal vPrimarySO As Double) As Double
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetQtyBySO_n_Product(vPrimarySO, vPrimaryProduct).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    ''' <summary>
    ''' mengembalikan item aktif booking GIT yang merujuk ke SO dan product tertentu
    ''' </summary>
    ''' <param name="vPrimarySO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJmlByPrimarySO_n_PrimaryProduct(ByVal vPrimarySO As Double, _
    ByVal vPrimaryProduct As String) As Integer
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitemTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetQtyBySO_n_Product(vPrimarySO, vPrimaryProduct).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByPrimarySO(ByVal vPrimarySO As Double) As Integer
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJmlByPrimarySO(vPrimarySO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetBookingNoByPrimarySO(ByVal vPrimarySO As Double, _
    ByVal vPrimaryProduct As String) As String
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invTableAdapter
        Dim TBL As New DseTransaksiBookingInv.tbltrans_booking_invDataTable
        Dim retval As String = ""
        Dim row As DseTransaksiBookingInv.tbltrans_booking_invRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_PrimarySO_n_PrimaryProduct(TBL, vPrimarySO, vPrimaryProduct) > 0 Then
                For Each row In TBL.Rows
                    If retval = "" Then
                        retval = row.No_Transaksi
                    Else
                        retval = retval + ", " + row.No_Transaksi
                    End If
                Next

                Return retval
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            TBL.Dispose()
            DTA.Dispose()
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal i As FrmTransaksiBookingInv) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'simpan data item
            For Each row In vdgvItem.Rows
                'simpan ke basis data item
                Dim ExpiryDate As Date
                Dim SecQty As Double
                Dim Batch As String = ""
                If row.Cells("ColExpiredDate").Value = "-" Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = row.Cells("ColExpiredDate").Value
                End If
                If String.IsNullOrEmpty(row.Cells("ColBatch").Value) = False Then
                    Batch = row.Cells("ColBatch").Value
                Else
                    Batch = ""
                End If
                If CStr(row.Cells("ColSecondaryQty").Value) = "-" Then
                    SecQty = -1
                Else
                    SecQty = row.Cells("ColSecondaryQty").Value
                End If
                DTA.Insert(vPrimaryTransaksi, row.Cells("ColPrimaryCodeNo").Value, row.Cells("ColQty").Value, _
                row.Cells("ColPrimaryUnits").Value, Batch, ExpiryDate, _
                row.Cells("ColRemark").Value, SecQty, row.Cells("ColPrimarySecondaryUnits").Value, _
                row.Cells("ColValidUntil").Value, row.Cells(i.ColPrimaryLocation.Index).Value, -1)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'hapus data item
            DTA.Delete(vPrimaryTransaksi)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
