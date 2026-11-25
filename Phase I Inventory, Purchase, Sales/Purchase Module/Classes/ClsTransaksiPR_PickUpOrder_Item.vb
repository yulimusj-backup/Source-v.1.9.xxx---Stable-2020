Option Explicit On
Option Strict Off
Public Class ClsTransaksiPR_PickUpOrder_Item
    Public Function SimpanItem(ByVal vPrimaryTransaksi As Double, _
                                ByVal vDGVItem As DataGridView, _
                                ByVal frm As FrmTransaksiPR_PickUpOrder, ByVal TBL_Dummy As DseTransaksiPR_PickUpOrder.TBLSetBalanceDataTable) As Boolean
        Dim DTAItem As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_order_itemTableAdapter
        Dim row As DataGridViewRow

        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim ObjUnit As New Core_Module.ClsMasterUnit
        Dim ObjPR As New ClsTransaksiPR

        Dim vPrimaryProduct As Double
        Dim vPrimaryUnit As Double
        Dim ListPrimaryPR As New List(Of Double)
        Dim row_Dummy As DseTransaksiPR_PickUpOrder.TBLSetBalanceRow

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For Each row In vDGVItem.Rows
                vDGVItem.AllowUserToAddRows = False
                vPrimaryProduct = ObjProduct.GetPrimary(row.Cells(frm.ColItem_CodeNo.Index).Value)
                vPrimaryUnit = ObjUnit.GetPrimary(row.Cells(frm.ColItem_Unit.Index).Value)

                With row
                    DTAItem.Insert(.Cells(frm.ColItem_PrimaryPR.Index).Value, vPrimaryTransaksi, _
                                    vPrimaryProduct, .Cells(frm.ColItem_Qty.Index).Value, _
                                    vPrimaryUnit, .Cells(frm.ColItem_Remarks.Index).Value)
                    If ListPrimaryPR.Contains(.Cells(frm.ColItem_PrimaryPR.Index).Value) = False Then
                        ListPrimaryPR.Add(.Cells(frm.ColItem_PrimaryPR.Index).Value)
                    End If
                End With
            Next

            'ambil primary pr, dari sebelum simpan perubahan, jika ada yang berbeda dgn dgv
            'maka set order balance juga
            For Each row_Dummy In TBL_Dummy.Rows
                If ListPrimaryPR.Contains(row_Dummy.Primary_PR) = False Then
                    ListPrimaryPR.Add(row_Dummy.Primary_PR)
                End If
            Next
            For Each PrimaryPR As Double In ListPrimaryPR
                ObjPR.SetOrderBalance1(PrimaryPR, ClsTransaksiPR.EnumReadyForPurchase.Yes)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            ListPrimaryPR = Nothing
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTAItem As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_order_itemTableAdapter
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Delete(vPrimaryTransaksi)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function IsiDGVItem(ByVal vPrimaryTransaksi As Double, ByVal vDGVItem As DataGridView, _
                              ByVal vFrm As FrmTransaksiPR_PickUpOrder, ByVal vObjCD As Core_Module.ClsCekData) As Boolean

        Dim DTAItem As New DseTransaksiPR_PickUpOrderTableAdapters.tbltrans_purchase_req_pickup_order_itemTableAdapter
        Dim TBLItem As New DseTransaksiPR_PickUpOrder.tbltrans_purchase_req_pickup_order_itemDataTable
        Dim rowItem As DseTransaksiPR_PickUpOrder.tbltrans_purchase_req_pickup_order_itemRow
        Dim rowDGV As DataGridViewRow

        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim ObjUnit As New Core_Module.ClsMasterUnit
        Dim ObjPR As New ClsTransaksiPR

        Try
            '''update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            '''autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            '''isikan data item
            vDGVItem.Rows.Clear()
            If DTAItem.FillBy_PrimaryTransaksi(TBLItem, vPrimaryTransaksi) > 0 Then
                For Each rowItem In TBLItem
                    vDGVItem.Rows.Add()
                    rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                    Dim DefPrec As Integer = ObjProduct.GetAngkaDecimal(rowItem.primary_code_no)
                    With rowDGV
                        .Cells(vFrm.ColItem_No.Index).Value = ""
                        .Cells(vFrm.ColItem_PrimaryPR.Index).Value = rowItem.primary_pr
                        .Cells(vFrm.ColItem_PRNo.Index).Value = ObjPR.GetNoTransaksi(rowItem.primary_pr)
                        .Cells(vFrm.ColItem_PrimaryCodeNo.Index).Value = rowItem.primary_code_no
                        .Cells(vFrm.ColItem_CodeNo.Index).Value = ObjProduct.GetProductCode(rowItem.primary_code_no)
                        .Cells(vFrm.ColItem_Desc.Index).Value = ObjProduct.GetDescription(rowItem.primary_code_no)
                        .Cells(vFrm.ColItem_Unit.Index).Value = ObjUnit.GetUnitCode(rowItem.primary_unit)
                        .Cells(vFrm.ColItem_Qty.Index).Value = ObjCD.DigitGrouping(rowItem.qty, DefPrec)
                        .Cells(vFrm.ColItem_Remarks.Index).Value = rowItem.remarks
                    End With
                Next
            End If
            '''setnumber
            vObjCD.SetNumber(vDGVItem)
            '''autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            '''nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem = Nothing
        End Try
    End Function
End Class
