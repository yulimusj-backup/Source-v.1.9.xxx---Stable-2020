Option Explicit On
Option Strict Off

Public Class ClsTransaksiPurchaseReturn_Balance_Import_Item

    ''isi item
    Public Function IsiDGVItem(ByVal vDGVItem As DataGridView, _
         ByVal Frm As FrmTransaksiPurchaseReturn_Balance_Import, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnBalance_ImportTableAdapters.tbltrans_purchasereturnbalance_import_itemTableAdapter
        Dim TBL As New DseTransaksiPurchaseReturnBalance_Import.tbltrans_purchasereturnbalance_import_itemDataTable
        Dim row As DseTransaksiPurchaseReturnBalance_Import.tbltrans_purchasereturnbalance_import_itemRow
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
                        rowDGV.Cells(Frm.ColPrimaryPR.Index).Value = .Primary_AOP
                        rowDGV.Cells(Frm.ColPrimarySupp.Index).Value = .Primary_Supp
                        rowDGV.Cells(Frm.ColPR_Number.Index).Value = .AOP_Number
                        rowDGV.Cells(Frm.ColPR_Date.Index).Value = Format(.AOP_Date, "dd MMM yyyy")
                        rowDGV.Cells(Frm.Col_PR_DueDate.Index).Value = Format(.AOP_Due_Date, "dd MMM yyyy")
                        rowDGV.Cells(Frm.ColSuppCode.Index).Value = objCust.GetSupplierCode(.Primary_Supp)
                        rowDGV.Cells(Frm.ColSuppName.Index).Value = objCust.GetSupplierName(.Primary_Supp)
                        rowDGV.Cells(Frm.ColPurchaseReturn.Index).Value = ObjCD.DigitGrouping(CStr(.Purchase))
                        rowDGV.Cells(Frm.ColVAT.Index).Value = ObjCD.DigitGrouping(CStr(.VAT))
                        rowDGV.Cells(Frm.ColDutystamp.Index).Value = ObjCD.DigitGrouping(CStr(.Dutystamp))
                        rowDGV.Cells(Frm.ColReceived.Index).Value = ObjCD.DigitGrouping(CStr(.Received))
                        rowDGV.Cells(Frm.ColBalance.Index).Value = ObjCD.DigitGrouping(CStr(.Balance))
                        rowDGV.Cells(Frm.ColTotal.Index).Value = ObjCD.DigitGrouping(CStr(.Total))
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
    ByVal Frm As FrmTransaksiPurchaseReturn_Balance_Import) As Boolean
        Dim row As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseTransaksiPurchaseReturnBalance_ImportTableAdapters.tbltrans_purchasereturnbalance_import_itemTableAdapter
        Dim DefaultPrecision As Integer = 0
        Dim DueDate As Date = Today
        Try
            'simpan data Produk Code No
            For Each row In vDGVItem.Rows
                If String.IsNullOrEmpty(row.Cells(Frm.ColRemarks.Index).Value.ToString) = True Then
                    row.Cells(Frm.ColRemarks.Index).Value = ""
                End If

                If row.Cells(Frm.ColPR_Number.Index).Value.ToString <> "" Then
                    'simpan ke basis data item
                    DTA.Insert(CDec(vPrimaryTransaksi), CDec(row.Cells(Frm.ColPrimaryPR.Index).Value), row.Cells(Frm.ColPR_Number.Index).Value.ToString, _
                     CDate(row.Cells(Frm.ColPR_Date.Index).Value), CDate(row.Cells(Frm.Col_PR_DueDate.Index).Value), CDec(row.Cells(Frm.ColPrimarySupp.Index).Value), _
                     CDec(row.Cells(Frm.ColPurchaseReturn.Index).Value), CDec(row.Cells(Frm.ColVAT.Index).Value), CDec(row.Cells(Frm.ColDutystamp.Index).Value), _
                     CDec(row.Cells(Frm.ColTotal.Index).Value), CDec(row.Cells(Frm.ColReceived.Index).Value), CDec(row.Cells(Frm.ColBalance.Index).Value), _
                     row.Cells(Frm.ColRemarks.Index).Value.ToString)
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
        Dim DTA As New DseTransaksiPurchaseReturnBalance_ImportTableAdapters.tbltrans_purchasereturnbalance_import_itemTableAdapter
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
