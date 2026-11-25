Option Explicit On
Option Strict Off

Public Class ClsTransaksiPackingListItem
    Public Function GetJml(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryPackingList As Double) As Double
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_itemTableAdapter

        Try
            Return DTA.ScalarQuery_GetJml(vPrimaryPackingList, vPrimaryCodeNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vPackingListNo As String) As Double
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimary(vPackingListNo).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary(vPackingListNo).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimaryPackingList As Double) As String
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglistTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_getNoTransaksi(vPrimaryPackingList)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_getNoTransaksi(vPrimaryPackingList)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_item_view1TableAdapter
        Dim TBL As New DseTransaksiPackingList.tbltrans_packinglist_item_view1DataTable
        Dim row As DseTransaksiPackingList.tbltrans_packinglist_item_view1Row

        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.Fill_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("0", row.Product_Code, row.Description, ._Supplier_s_Code_No, ._Supplier_s_Desc, _
                        row.Qty_Master, row.Qty_Per_Master, row.Master_Unit, row.Qty, .Total_Qty, .Unit_Code, row.Remarks, _
                        .Primary_Product, row.Primary_Units, row.Primary_Master_Units, .Primary_PO)
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
    Public Function IsiDgvCheckStock(ByVal vDGV As DataGridView, _
     ByVal TBL As DseTransaksiPackingListItem.DataTable2DataTable, _
     Optional ByVal vPrimaryProduct As Double = -1) As Boolean
        Dim DTA As New DseTransaksiPackingListItemTableAdapters.DataTable2TableAdapter
        Try
            If vPrimaryProduct = -1 Then
                DTA.Fill(TBL)
            Else
                DTA.FillBy_PrimaryProduct(TBL, vPrimaryProduct)
            End If
            vDGV.DataSource = TBL.DefaultView
            vDGV.Columns(1).DefaultCellStyle.Format = "dd MMM yyyy"
            vDGV.Columns(4).DefaultCellStyle.Format = "dd MMM yyyy"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function Simpan(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_itemTableAdapter

        Try
            For Each row In vdgvItem.Rows
                'simpan ke basis data item
                DTA.Insert(vPrimaryTransaksi, row.Cells(13).Value, row.Cells(16).Value, row.Cells(6).Value, _
                row.Cells(7).Value, row.Cells(15).Value, row.Cells(9).Value, _
                row.Cells(10).Value, row.Cells(14).Value, row.Cells(12).Value)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_itemTableAdapter

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

    ''' <summary>
    ''' fungsi ini digunakan untuk cek stock. Mengembalikan total barang yang primarynya diberikan
    ''' Untuk GIT yang belum dimasukkan ke dalam stock in
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJml(ByVal vPrimaryCodeNo As Double) As Double
        Dim DTA As New DseTransaksiPackingListItemTableAdapters.tbltrans_packinglist_itemTableAdapter
        Try
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
End Class
