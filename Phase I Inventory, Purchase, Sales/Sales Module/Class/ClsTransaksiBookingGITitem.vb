Option Explicit On
Option Strict Off

Public Class ClsTransaksiBookingGITitem
    Public Sub IsiTBLBooking(ByVal vPrimaryBooking As Double, _
    ByVal TBL As DseTransaksiBookingGIT.DataTable3DataTable)
        Dim DTA As New DseTransaksiBookingGITTableAdapters.DataTable3TableAdapter

        Try
            DTA.FillBy_PrimaryBooking(TBL, vPrimaryBooking)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLBooking2(ByVal vPrimaryBooking As Double, _
    ByVal TBL As DseTransaksiBookingGIT.DataTable3DataTable, ByVal vPrimarySO As Double)
        Dim DTA As New DseTransaksiBookingGITTableAdapters.DataTable3TableAdapter

        Try
            DTA.FillBy_PrimaryBooking_PrimarySO(TBL, vPrimaryBooking, vPrimarySO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' fungsi ini digunakan untuk cek stock. Mengembalikan total barang yang primarynya diberikan
    ''' Untuk booking GIT yang masih aktif
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJml_BookingGIT_Aktif(ByVal vPrimaryCodeNo As Double) As Double
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_bookinggit_item_aktif_viewTableAdapter

        Try
            Return DTA.ScalarQuery_GetJml(vPrimaryCodeNo).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlBooking(ByVal vPrimaryProduct As Double, ByVal vPrimarySO As Double) As Double
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_gititem2TableAdapter
        Try
            Return DTA.ScalarQuery_GetJml(vPrimarySO, vPrimaryProduct).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDgvCheckStock(ByVal vdgV As DataGridView, _
    ByVal TBL As DseTransaksiBookingGIT.tbltrans_bookinggit_item_aktif_viewDataTable, _
    Optional ByVal VPrimaryProduct As Double = -1) As Boolean
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_bookinggit_item_aktif_viewTableAdapter

        Try
            If VPrimaryProduct = -1 Then
                DTA.Fill(TBL)
            Else
                DTA.FillBy_PrimaryProduct(TBL, VPrimaryProduct)
            End If
            vdgV.DataSource = TBL.DefaultView
            vdgV.Columns(1).DefaultCellStyle.Format = "dd MMM yyyy"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan item aktif booking GIT yang merujuk ke SO tertentu
    ''' </summary>
    ''' <param name="vPrimarySO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJmlTransaksi_ByPrimarySONPrimaryProduct(ByVal vPrimarySO As Double, _
    ByVal vPrimaryProduct As Double) As Integer
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_gitTableAdapter

        Try
            Return DTA.ScalarQuery_GetJml(vPrimarySO, vPrimaryProduct).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetQtyByPrimarySO_n_Product_Default(ByVal vPrimarySO As Double, _
    ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_bookinggit_item_aktif_viewTableAdapter

        Try
            Return DTA.ScalarQuery_GetQtyByPrimarySO_n_Product_Default(vPrimaryProduct, vPrimarySO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetQtyByPrimarySO_n_Product_Multi(ByVal vPrimarySO As Double, _
    ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_bookinggit_item_aktif_viewTableAdapter

        Try
            Return DTA.ScalarQuery_GetQtyByPrimarySO_n_Product_Multi(vPrimaryProduct, vPrimarySO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetBookingNoByPrimarySO_n_Product(ByVal vPrimarySO As Double, _
    ByVal vPrimaryProduct As String) As String
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_gitTableAdapter
        Dim TBL As New DseTransaksiBookingGIT.tbltrans_booking_gitDataTable
        Dim retval As String = ""
        Dim row As DseTransaksiBookingGIT.tbltrans_booking_gitRow

        Try
            If DTA.Fill(TBL, vPrimarySO, vPrimaryProduct) > 0 Then
                For Each row In TBL.Rows
                    If retval = "" Then
                        retval = row.Booking_No
                    Else
                        retval = retval + ", " + row.Booking_No
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
    Public Sub IsikanTBLSO(ByVal vPrimaryCustomer As Double, ByVal TBL As DseTransaksiBookingGIT.DataTable22DataTable)
        Dim DTA As New DseTransaksiBookingGITTableAdapters.DataTable22TableAdapter
        Try
            DTA.FillBy_PrimaryCustomer(TBL, vPrimaryCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
    Public Function IsiTBLItem(ByVal vPrimarySO As Double, _
    ByVal TBL As DseTransaksiBookingGIT.tbltrans_booking_GITitem1DataTable) As Boolean
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_GITitem1TableAdapter
        Try
            TBL.Rows.Clear()
            DTA.Fill(TBL, vPrimarySO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Sub IsikanTBLProductBySO(ByVal vPrimaryBooking As Double, ByVal TBL As DseTransaksiBookingGIT.DataTable22DataTable, _
       ByVal vPrimarySO As Double)
        Dim DTA As New DseTransaksiBookingGITTableAdapters.DataTable22TableAdapter
        Try
            DTA.FillBy_PrimarySONBooking(TBL, vPrimarySO, vPrimaryBooking)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
    Public Sub IsiTBLBatchNo(ByVal vPrimaryBooking As Double, ByVal vPrimaryProduct As Double, _
    ByVal TBL As DseTransaksiBookingGIT.tbltrans_do_item1DataTable)
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_do_item1TableAdapter
        Try
            DTA.Fill(TBL, vPrimaryProduct, vPrimaryBooking)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
    Public Sub IsiSecondaryQty(ByVal vPrimaryBooking As Double, ByVal vPrimaryProduct As Double, _
    ByVal vBatchNo As String, ByVal TBL As DseTransaksiBookingGIT.tbltrans_do_item2DataTable)
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_do_item2TableAdapter
        Try
            DTA.Fill(TBL, vPrimaryBooking, vPrimaryProduct, vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' fungsi ini digunakan untuk cek stock. Mengembalikan total barang yang primarynya diberikan
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJmlByPrimaryProductNBooking(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryBooking As Double) As Double
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_gititemTableAdapter
        Try
            If DTA.ScalarQuery_getJmlByPrimaryProductNPrimaryBooking(vPrimaryCodeNo, vPrimaryBooking).HasValue = True Then
                Return DTA.ScalarQuery_getJmlByPrimaryProductNPrimaryBooking(vPrimaryCodeNo, vPrimaryBooking).Value
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
    Public Function GetJmlByPrimaryProductNBookingNPrimarySO(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryBooking As Double, ByVal vPrimarySO As Double) As Double
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_gititemTableAdapter
        Try
            If DTA.ScalarQuery_getJmlByPrimaryProductNPrimaryBookingNPrimarySO(vPrimaryCodeNo, vPrimaryBooking, vPrimarySO).HasValue = True Then
                Return DTA.ScalarQuery_getJmlByPrimaryProductNPrimaryBookingNPrimarySO(vPrimaryCodeNo, vPrimaryBooking, vPrimarySO).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlOtherBookingGIT2(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryBooking As Double) As Double
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_bookinggit_item_aktif_viewTableAdapter

        Try
            Return DTA.ScalarQuery_GetJml2(vPrimaryCodeNo, vPrimaryBooking).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiBookingGITTableAdapters.tbltrans_booking_gititemTableAdapter
        Try
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
