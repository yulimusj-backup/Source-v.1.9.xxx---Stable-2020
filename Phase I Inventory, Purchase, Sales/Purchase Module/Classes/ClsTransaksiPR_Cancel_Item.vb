Option Explicit On
Option Strict Off

Public Class ClsTransaksiPR_Cancel_Item
    Public Function SimpanItem(ByVal vPrimaryTransaksi As Double, _
                                    ByVal vDGVItem As DataGridView, _
                                    ByVal frm As FrmTransaksiPR_Cancel) As Boolean

        Dim DTAItem As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancel_itemTableAdapter
        Dim row As DataGridViewRow

        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim ObjUnit As New Core_Module.ClsMasterUnit
        Dim ObjPR As New ClsTransaksiPR

        Dim vPrimaryProduct As Double
        Dim vPrimaryUnit As Double

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            'vDGVItem.AllowUserToAddRows = False

            For Each row In vDGVItem.Rows
                vDGVItem.AllowUserToAddRows = False
                vPrimaryProduct = ObjProduct.GetPrimary(row.Cells(frm.ColItem_CodeNo.Index).Value)
                vPrimaryUnit = ObjUnit.GetPrimary(row.Cells(frm.ColItem_Unit.Index).Value)
                With row
                    DTAItem.Insert(vPrimaryTransaksi, vPrimaryProduct, _
                                    .Cells(frm.ColItem_Qty.Index).Value, vPrimaryUnit, _
                                    .Cells(frm.ColItem_Remarks.Index).Value)
                End With

            Next
            ObjPR.SetOrderBalance1(frm.TxtPrimaryPR.Text, ClsTransaksiPR.EnumReadyForPurchase.Yes)
            vDGVItem.AllowUserToAddRows = True
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryTransaksi As Double, _
                              ByVal frm As FrmTransaksiPR_Cancel) As Boolean

        Dim DTAItem As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancel_itemTableAdapter
        Dim ObjPR As New ClsTransaksiPR
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
                                ByVal i As FrmTransaksiPR_Cancel, ByVal vObjCD As Core_Module.ClsCekData) As Boolean

        Dim DTAItem As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancel_itemTableAdapter
        Dim TBLItem As New DseTransaksiPR_Cancel.tbltrans_purchase_req_cancel_itemDataTable
        Dim rowItem As DseTransaksiPR_Cancel.tbltrans_purchase_req_cancel_itemRow
        Dim rowDGV As DataGridViewRow

        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim ObjUnit As New Core_Module.ClsMasterUnit
        Dim ObjPR As New ClsTransaksiPR

        Try
            '   update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            '   autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            '   isikan data item
            vDGVItem.Rows.Clear()

            If DTAItem.FillBy_Primary(TBLItem, vPrimaryTransaksi) > 0 Then
                For Each rowItem In TBLItem
                    vDGVItem.Rows.Add()
                    vDGVItem.AllowUserToAddRows = False
                    rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)

                    Dim DefPrec As Integer = ObjProduct.GetAngkaDecimal(rowItem.primary_code_no)
                    With rowDGV
                        .Cells(i.ColItem_No.Index).Value = ""
                        .Cells(i.ColItem_PrimaryCode.Index).Value = rowItem.primary_code_no
                        .Cells(i.ColItem_CodeNo.Index).Value = ObjProduct.GetProductCode(rowItem.primary_code_no)
                        .Cells(i.ColItem_Desc.Index).Value = ObjProduct.GetDescription(rowItem.primary_code_no)
                        .Cells(i.ColItem_Unit.Index).Value = ObjUnit.GetUnitCode(rowItem.primary_unit)
                        .Cells(i.ColItem_Qty.Index).Value = ObjCD.DigitGrouping(rowItem.qty, DefPrec)
                        .Cells(i.ColItem_Remarks.Index).Value = rowItem.remarks
                    End With
                Next
            End If

            vDGVItem.AllowUserToAddRows = False

            '   setnumber
            vObjCD.SetNumber(vDGVItem)

            '   autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            '   nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem = Nothing
        End Try
    End Function
End Class
