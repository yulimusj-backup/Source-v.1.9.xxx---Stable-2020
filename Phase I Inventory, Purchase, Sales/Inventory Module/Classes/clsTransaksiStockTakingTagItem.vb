Option Explicit On
Option Strict Off

Public Class clsTransaksiStockTakingTagItem
    Public Function IsiDGV(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiStockTagTableAdapters.trans_stocktaking_tag_item_viewTableAdapter
        Dim TBL As New DseTransaksiStockTag.trans_stocktaking_tag_item_viewDataTable
        Dim row As DseTransaksiStockTag.trans_stocktaking_tag_item_viewRow
        Dim Sec_Units As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    If row.Primary_Secondary_Units <> -1 Then
                        Sec_Units = row.Secondary_Unit_Code
                    Else
                        Sec_Units = ""
                    End If

                    vDGVItem.Rows.Add("0", row.Product_Code, row.Description, _
                    ObjCD.DigitGrouping(row.Qty, row._Decimal), row.Unit_Code, ObjCD.DigitGrouping(row.Secondary_Qty, row._Decimal), Sec_Units, row.Batch_No, row.Expiry_Date, _
                    row.Remarks, row.Primary_Product, row.Primary_Units, row.Primary_Secondary_Units)
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

    Public Function Hapus(ByVal vPrimaryStockTag As Double) As Boolean
        Dim DTA As New DseTransaksiStockTagTableAdapters.tbltrans_stock_taking_tag_itemTableAdapter

        Try
            DTA.Delete(vPrimaryStockTag)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItem As DataGridView, ByVal vTanggalTransaksi As Date, _
    ByVal vPrimaryTransaksi As Double, ByVal vNoTransaksi As String) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiStockTagTableAdapters.tbltrans_stock_taking_tag_itemTableAdapter

        Dim Sec_Qty As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#

        Try
            For Each row In vdgvItem.Rows
                BatchNo = row.Cells(7).Value
                If IsNumeric(row.Cells(5).Value) = True Then
                    Sec_Qty = row.Cells(5).Value
                Else
                    Sec_Qty = "-1"
                End If
                If IsDate(row.Cells(8).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells(8).Value)
                End If

                'simpan ke basis data item
                DTA.Insert(vPrimaryTransaksi, row.Cells(10).Value, row.Cells(3).Value, row.Cells(11).Value, _
                Sec_Qty, row.Cells(12).Value, BatchNo, ExpiryDate, row.Cells(9).Value)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
