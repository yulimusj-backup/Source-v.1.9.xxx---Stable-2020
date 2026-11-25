Option Explicit On
Option Strict Off

Public Class ClsTransaksiGRNItem

    Public Function GetJmlByPO(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double) As Integer
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grn_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJmlBYPO_n_Product(vPrimaryPO, vPrimaryProduct).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByPO(ByVal vPrimaryPO As Double) As Integer
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grn_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJmlByPrimaryPO(vPrimaryPO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grn_itemTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjLS As New Core_Module.clsMasterListedSoftware
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim IsInventoryInstalled As Boolean = ObjLS.IsInventoryInstalled
        Dim Sec_Qty As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim objPO As New ClsTransaksiPO
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Core_Module.SetConstring_DTAProductQty(DTA1)

            For Each row In vdgvItem.Rows
                BatchNo = row.Cells("ColBatchNo").Value
                If IsNumeric(row.Cells("ColSecQty").Value) = True Then
                    Sec_Qty = row.Cells("ColSecQty").Value
                Else
                    Sec_Qty = "-1"
                End If

                If IsDate(row.Cells.Item("ColExpiryDate").Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells.Item("ColExpiryDate").Value)
                End If
                'If IsInventoryInstalled = True Then
                'AuthorizedLevelNeeded = 1
                'End If

                'simpan ke basis data item
                DTA.Insert(vPrimaryTransaksi, row.Cells("ColPrimaryCodeNo").Value, row.Cells("colPrimaryPO").Value, row.Cells("ColBatchNo").Value, _
                ExpiryDate, row.Cells("ColQtyMaster").Value, row.Cells("ColPrimaryMasterUnits").Value, _
                row.Cells("ColQty").Value, _
                row.Cells("ColTotalQty").Value, row.Cells("ColPrimaryUnit").Value, Sec_Qty, _
                row.Cells("ColPrimarySecUnit").Value, row.Cells("DataGridViewTextBoxColumn13").Value, row.Cells("ColQtyPerMaster").Value)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjProd = Nothing
            ObjLS = Nothing
            objPO = Nothing
        End Try
    End Function

    Public Function SimpanLewatTransaksiStockInOut_Default(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grn_itemTableAdapter
        Dim Sec_Qty As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each row In vdgvItem.Rows
                BatchNo = row.Cells("ColBatch").Value
                If IsNumeric(row.Cells("ColSecondaryQty").Value) = True Then
                    Sec_Qty = row.Cells("ColSecondaryQty").Value
                Else
                    Sec_Qty = "-1"
                End If

                If IsDate(row.Cells.Item("ColExpiredDate").Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells.Item("ColExpiredDate").Value)
                End If
                'If IsInventoryInstalled = True Then
                'AuthorizedLevelNeeded = 1
                'End If

                'simpan ke basis data item
                DTA.Insert(vPrimaryTransaksi, row.Cells("ColPrimaryCodeNo").Value, row.Cells("ColPrimaryPOno").Value, _
                row.Cells("ColBatch").Value, ExpiryDate, 0, -1, row.Cells("ColQty").Value, _
                row.Cells("ColQty").Value, row.Cells("ColPrimaryUnits").Value, Sec_Qty, _
                row.Cells("ColPrimarySecondaryUnits").Value, row.Cells("ColRemark").Value, 0)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function SimpanLewatTransaksiStockInOut_Multi(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grn_itemTableAdapter
        Dim Sec_Qty As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#

        Dim vFrmRef As New FrmTransaksiStockInPurchase2
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each row In vdgvItem.Rows
                BatchNo = row.Cells("ColBatch").Value
                If IsNumeric(row.Cells("ColSecondaryQty").Value) = True Then
                    Sec_Qty = row.Cells("ColSecondaryQty").Value
                Else
                    Sec_Qty = "-1"
                End If

                If IsDate(row.Cells.Item("ColExpiredDate").Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells.Item("ColExpiredDate").Value)
                End If
                'If IsInventoryInstalled = True Then
                'AuthorizedLevelNeeded = 1
                'End If

                'simpan ke basis data item
                DTA.InsertQuery(vPrimaryTransaksi, row.Cells("ColPrimaryCodeNo").Value, row.Cells("ColPrimaryPOno").Value, _
                row.Cells("ColBatch").Value, ExpiryDate, 0, -1, row.Cells("ColQty").Value, _
                row.Cells("ColQty").Value, row.Cells("ColPrimaryUnits").Value, Sec_Qty, _
                row.Cells("ColPrimarySecondaryUnits").Value, row.Cells("ColRemark").Value, 0, _
                row.Cells(vFrmRef.ColQtyMulti.Index).Value, row.Cells(vFrmRef.ColPrimaryUnitMulti.Index).Value, _
                row.Cells(vFrmRef.ColRatioReference.Index).Value, row.Cells(vFrmRef.ColRatioDefault.Index).Value)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grn_itemTableAdapter
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
        Dim DTA As New DseTransaksiGRNTableAdapters.TblDGV_ItemTableAdapter
        Dim TBL As New DseTransaksiGRN.TblDGV_ItemDataTable
        Dim row As DseTransaksiGRN.TblDGV_ItemRow
        Dim Sec_Units As String = ""
        Dim ObjCD As New Core_Module.ClsCekData
        Dim objpo As New ClsTransaksiPO
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.FillBy_PrimaryGRN(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    If row.Primary_Secondary_Units <> -1 Then
                        Sec_Units = row.SecUnit
                    Else
                        Sec_Units = ""
                    End If

                    With row
                        vDGVItem.Rows.Add("0", row.Product_Code, row.ProductDesc, row.Supplier_Code, "", row.NoPO, _
                        row.Batch_No, row.Expiry_Date, row.Qty_Master, row.Qty_Per_Master, .MasterUnit, row.Qty, _
                        .Total_Qty, row.unit, row.Secondary_Qty, Sec_Units, row.Remarks, row.Primary_Product, _
                        row.Primary_Master_Units, row.Primary_Units, row.Primary_Secondary_Units, row.Primary_PO)
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
End Class
