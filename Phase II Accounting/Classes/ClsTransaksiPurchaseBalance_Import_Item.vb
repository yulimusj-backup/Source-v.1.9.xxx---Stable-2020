Option Explicit On
Option Strict Off

Public Class ClsTransaksiPurchaseBalance_Import_Item
    ''isi item
    Public Function IsiDGVItem(ByVal vDGVItem As DataGridView, _
         ByVal Frm As FrmTransaksiPurchaseBalance_Import, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiPurchaseBalance_ImportTableAdapters.tbltrans_purchasebalance_import_itemTableAdapter
        Dim TBL As New DseTransaksiPurchaseBalance_Import.tbltrans_purchasebalance_import_itemDataTable
        Dim row As DseTransaksiPurchaseBalance_Import.tbltrans_purchasebalance_import_itemRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim objCust As New Core_Module.clsMasterSupplier
        Dim rowDGV As DataGridViewRow
        Dim DefaultPrecision As Integer = 0
        Try
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()

            'isikan tabel
            If DTA.FillBy_PrimaryTransaksi(TBL, CDec(vPrimaryTransaksi)) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColNo.Index).Value = ""
                        rowDGV.Cells(Frm.ColPrimaryAOP.Index).Value = .Primary_AOP
                        rowDGV.Cells(Frm.ColPrimarySupp.Index).Value = .Primary_Supp
                        rowDGV.Cells(Frm.ColAOP_Number.Index).Value = .AOP_Number
                        rowDGV.Cells(Frm.ColAOP_Date.Index).Value = Format(.AOP_Date, "dd MMM yyyy")
                        rowDGV.Cells(Frm.Col_AOP_DueDate.Index).Value = Format(.AOP_Due_Date, "dd MMM yyyy")
                        rowDGV.Cells(Frm.ColSuppCode.Index).Value = objCust.GetSupplierCode(.Primary_Supp)
                        rowDGV.Cells(Frm.ColSuppName.Index).Value = objCust.GetSupplierName(.Primary_Supp)
                        rowDGV.Cells(Frm.ColPurchase.Index).Value = ObjCD.DigitGrouping(.Purchase.ToString)
                        rowDGV.Cells(Frm.ColVAT.Index).Value = ObjCD.DigitGrouping(.VAT.ToString)
                        rowDGV.Cells(Frm.ColDutystamp.Index).Value = ObjCD.DigitGrouping(.Dutystamp.ToString)
                        rowDGV.Cells(Frm.ColCost.Index).Value = ObjCD.DigitGrouping(.Cost.ToString)
                        rowDGV.Cells(Frm.ColPaid.Index).Value = ObjCD.DigitGrouping(.Received.ToString)
                        rowDGV.Cells(Frm.ColBalance.Index).Value = ObjCD.DigitGrouping(.Balance.ToString)
                        rowDGV.Cells(Frm.ColTotal.Index).Value = ObjCD.DigitGrouping(.Total.ToString)
                        rowDGV.Cells(Frm.ColRemarks.Index).Value = .Remarks
                    End With
                Next
            End If
            'set nomer
            ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            objCust = Nothing
        End Try
    End Function

    Public Function SimpanItem(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTransaksiPurchaseBalance_Import) As Boolean
        Dim row As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseTransaksiPurchaseBalance_ImportTableAdapters.tbltrans_purchasebalance_import_itemTableAdapter
        Dim DefaultPrecision As Integer = 0
        Dim DueDate As Date = Today
        Try
            'simpan data Produk Code No
            For Each row In vDGVItem.Rows
                If String.IsNullOrEmpty(row.Cells(Frm.ColRemarks.Index).Value.ToString) = True Then
                    row.Cells(Frm.ColRemarks.Index).Value = ""
                End If

                If row.Cells(CInt(Frm.ColAOP_Number.Index)).Value.ToString <> "" Then
                    'simpan ke basis data item
                    DTA.Insert(CDec(vPrimaryTransaksi), CDec(row.Cells(Frm.ColPrimaryAOP.Index).Value), row.Cells(Frm.ColAOP_Number.Index).Value.ToString, _
                     CDate(row.Cells(Frm.ColAOP_Date.Index).Value), CDate(row.Cells(Frm.Col_AOP_DueDate.Index).Value), CDec(row.Cells(Frm.ColPrimarySupp.Index).Value), _
                     CDec(row.Cells(Frm.ColPurchase.Index).Value), CDec(row.Cells(Frm.ColVAT.Index).Value), CDec(row.Cells(Frm.ColDutystamp.Index).Value), _
                     CDec(row.Cells(Frm.ColCost.Index).Value), CDec(row.Cells(Frm.ColTotal.Index).Value), CDec(row.Cells(Frm.ColPaid.Index).Value), _
                     CDec(row.Cells(Frm.ColBalance.Index).Value), row.Cells(Frm.ColRemarks.Index).Value.ToString)
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiPurchaseBalance_ImportTableAdapters.tbltrans_purchasebalance_import_itemTableAdapter
        Try
            'hapus data item
            DTA.Delete(CDec(vPrimaryTransaksi))
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
