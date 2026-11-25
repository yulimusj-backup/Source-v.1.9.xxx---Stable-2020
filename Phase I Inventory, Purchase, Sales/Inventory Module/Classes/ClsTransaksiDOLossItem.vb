Option Explicit On
Option Strict Off

Public Class ClsTransaksiDOLossItem

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_loss_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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

    Public Function IsiDGV(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_dolossview2TableAdapter
        Dim TBL As New DseTransaksiDOLoss.tbltrans_dolossview2DataTable
        Dim row As DseTransaksiDOLoss.tbltrans_dolossview2Row
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.FillBy_DoLossNo(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("0", .Product_Code, .Description, ._Customer_sCode_, ._Customer_sDesc_, .NoSO, _
                        .Batch_No, .Expiry_Date, .Qty_Master, .Qty_Per_Master, .Master_Unit, .Qty, .Total_Qty, .Units, .Secondary_Qty, _
                        .SecUnit, .Remarks, .Primary_Product, .Primary_SO, .Primary_Master_Units, .Primary_Units, .Primary_Secondary_Units)
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

    Public Sub IsikanTBLSO(ByVal vPrimaryDOLoss As Double, ByVal TBL As DseTransaksiDOLoss.tbltrans_doloss_item_view1DataTable)
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_doloss_item_view1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryDOLoss(TBL, vPrimaryDOLoss)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLBatchNo(ByVal vPrimaryDOLoss As Double, ByVal vPrimaryProduct As Double, _
    ByVal TBL As DseTransaksiDOLoss.tbltrans_do_loss_item2DataTable)
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_loss_item2TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryProduct, vPrimaryDOLoss)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLProductBySO(ByVal vPrimaryDOLoss As Double, ByVal TBL As DseTransaksiDOLoss.tbltrans_doloss_item_view1DataTable, _
    ByVal vPrimarySO As Double)
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_doloss_item_view1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimarySO_PrimaryDOLoss(TBL, vPrimaryDOLoss, vPrimarySO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetJmlPendingDO(ByVal vPrimaryCodeNo As Double, ByVal vPrimarySO As Double) As Double
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_loss_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetPendingDO2(vPrimaryCodeNo, vPrimarySO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlPendingDO(ByVal vPrimaryCodeNo As Double) As Double
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_loss_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetPendingDO(vPrimaryCodeNo).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItem As DataGridView, _
   ByVal vPrimaryTransaksi As Double) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_loss_itemTableAdapter
        Dim ExpiryDate As Date = #1/1/1950#
        Dim SecQTy As Double = -1

        'Dim DelDate As Date
        Dim PriceCondition As Integer = 0
        Dim objProduct As New Core_Module.clsMasterProduct
        Dim objitem As New ClsTransaksiDOLossItem
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each row In vdgvItem.Rows
                'simpan ke basis data item
                If IsNumeric(row.Cells(14).Value) = Boolean.TrueString Then
                    SecQTy = row.Cells(14).Value
                Else
                    SecQTy = -1
                End If
                If IsDate(row.Cells(7).Value) = True Then
                    ExpiryDate = row.Cells(7).Value
                Else
                    ExpiryDate = #1/1/1950#
                End If

                DTA.Insert(vPrimaryTransaksi, row.Cells(17).Value, row.Cells(18).Value, row.Cells(6).Value, _
                row.Cells(8).Value, row.Cells(19).Value, row.Cells(11).Value, row.Cells(12).Value, row.Cells(20).Value, _
                SecQTy, row.Cells(21).Value, row.Cells(16).Value, row.Cells(9).Value, ExpiryDate)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            objProduct = Nothing
            objitem = Nothing
        End Try
    End Function

    Public Sub IsiSecondaryQty(ByVal vPrimaryDOLoss As Double, ByVal vPrimaryProduct As Double, _
    ByVal vBatchNo As String, ByVal TBL As DseTransaksiDOLoss.tbltrans_do_loss_item3DataTable)
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_loss_item3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryDOLoss, vPrimaryProduct, vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetJml(ByVal vPrimaryDOLoss As Double, ByVal vPrimaryCodeNo As Double, _
    ByVal SecQTy As Double, ByVal BatchNo As String) As Double
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_loss_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJml2(vPrimaryDOLoss, vPrimaryCodeNo, BatchNo, SecQTy).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryDOLoss As Double) As Double
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_loss_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJml1(vPrimaryCodeNo, vPrimaryDOLoss).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vPrimaryTransaksi As Double, ByVal vPrimarySO As Double, _
    ByVal vPrimaryCodeNo As Double) As Double
        Dim DTA As New DseTransaksiDOLossTableAdapters.tbltrans_do_loss_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJml(vPrimaryTransaksi, vPrimaryCodeNo, vPrimarySO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
