Option Explicit On
Option Strict Off
Public Class ClsTransaksiBarcode
    Public Enum EnumInputType As Integer
        Add
        Replace
    End Enum

    Public Sub GetDataInput(ByRef vTBL As Core_Module.DseBarcodeInput.DataTable1DataTable, ByRef vInputType As EnumInputType, _
    ByVal vShowTargetLocation As Boolean, ByVal vTipeForm As FrmTransaksiStockAdvanceAdd.EnumRefDoc, _
    ByVal vDgv As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String)
        Dim ij As New FrmTransaksiStockAdvanceAdd
        Dim tblDgv As New Core_Module.DseBarcodeInput.tbl_qty_itemDataTable

        Try
            vTBL.Rows.Clear()

            'menentukan menggunakan location dan warehouse target
            If vShowTargetLocation = False Then
                ij.TxtLocationTarget.Enabled = False
                ij.BtnFindLocationTarget.Enabled = False
            End If

            Me.IsikanTblDgv(vTipeForm, vDgv, tblDgv)

            ij.ParentForm1 = vTipeForm
            ij.vTBL = vTBL
            ij.vTblDgv = tblDgv
            ij.vPrimaryTransaksi = vPrimaryTransaksi
            ij.UserCode = vUserCode

            If ij.ShowDialog() <> DialogResult.OK Then
                If ij.vTBL.Rows.Count > 0 Then
                    ij.vTBL.Rows.Clear()
                End If
            End If

            'tipe input add/replace
            vInputType = ij.InputType
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub IsikanTblDgv(ByVal vTipeForm As FrmTransaksiStockAdvanceAdd.EnumRefDoc, ByVal vDgv As DataGridView, _
    ByRef vTBL As Core_Module.DseBarcodeInput.tbl_qty_itemDataTable)
        Dim row As Core_Module.DseBarcodeInput.tbl_qty_itemRow
        Dim PrimaryProductIdx As Integer = -1
        Dim PrimaryLocIdx As Integer = -1
        Dim BatchIdx As Integer = -1
        Dim QtyIdx As Integer = -1

        Try
            If vTipeForm = FrmTransaksiStockAdvanceAdd.EnumRefDoc.Location_Transfer_Source Then
                Dim ij As New FrmTransaksiStockTransfer
                PrimaryProductIdx = ij.ColPrimaryCodeNo.Index
                PrimaryLocIdx = ij.ColPrimaryLocation.Index
                BatchIdx = ij.ColBatch.Index
                QtyIdx = ij.ColQty.Index
                ij.Dispose()
            ElseIf vTipeForm = FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_Bundling_Item_Before Then
                Dim ij As New FrmTransaksiStockBundling
                PrimaryProductIdx = ij.ColPrimaryCodeNo.Index
                PrimaryLocIdx = ij.ColPrimaryLocation.Index
                BatchIdx = ij.ColBatch.Index
                QtyIdx = ij.ColQty.Index
                ij.Dispose()
            ElseIf vTipeForm = FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_Bundling_Item_After Then
                Dim ij As New FrmTransaksiStockBundling
                PrimaryProductIdx = ij.ColPrimaryCodeNoAfter.Index
                PrimaryLocIdx = ij.ColPrimaryLocationAfter.Index
                BatchIdx = ij.ColBatchNoAfter.Index
                QtyIdx = ij.ColQtyAfter.Index
                ij.Dispose()
            ElseIf vTipeForm = FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_In_Adjust Then
                Dim ij As New FrmTransaksiStockInAdj
                PrimaryProductIdx = ij.ColPrimaryCodeNo.Index
                PrimaryLocIdx = ij.ColPrimaryLocation.Index
                BatchIdx = ij.ColBatch.Index
                QtyIdx = ij.ColQty.Index
                ij.Dispose()
            ElseIf vTipeForm = FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_Out_Adjust Then
                Dim ij As New FrmTransaksiStockOutAdj
                PrimaryProductIdx = ij.ColPrimaryCodeNo.Index
                PrimaryLocIdx = ij.ColPrimaryLocation.Index
                BatchIdx = ij.ColBatch.Index
                QtyIdx = ij.ColQty.Index
                ij.Dispose()
            End If

            For Each rd As DataGridViewRow In vDgv.Rows
                row = vTBL.Newtbl_qty_itemRow
                row.Primary_Product = CDbl(rd.Cells(PrimaryProductIdx).Value)
                row.Primary_Location = CDbl(rd.Cells(PrimaryLocIdx).Value)
                row.Batch_No = rd.Cells(BatchIdx).Value
                row.Qty = CDec(rd.Cells(QtyIdx).Value)
                vTBL.Rows.Add(row)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
