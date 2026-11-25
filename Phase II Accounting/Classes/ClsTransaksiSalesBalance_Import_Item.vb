Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesBalance_Import_Item
    ''isi item
    Public Function IsiDGVItem(ByVal vDGVItem As DataGridView, _
         ByVal Frm As FrmTransaksiSalesBalance_Import, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiSalesBalance_ImportTableAdapters.tbltrans_salesbalance_import_itemTableAdapter
        Dim TBL As New DseTransaksiSalesBalance_Import.tbltrans_salesbalance_import_itemDataTable
        Dim row As DseTransaksiSalesBalance_Import.tbltrans_salesbalance_import_itemRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim objCust As New Core_Module.clsMasterCustomer
        Dim rowDGV As DataGridViewRow
        Dim DefaultPrecision As Integer = 0
        Try
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()

            'isikan tabel
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColNo.Index).Value = ""
                        rowDGV.Cells(Frm.ColPrimaryInv.Index).Value = .Primary_Invoice
                        rowDGV.Cells(Frm.ColPrimaryCust.Index).Value = .Primary_Cust
                        rowDGV.Cells(Frm.ColInv_Number.Index).Value = .Inv_Number
                        rowDGV.Cells(Frm.ColInv_Date.Index).Value = Format(.Inv_Date, "dd MMM yyyy")
                        rowDGV.Cells(Frm.Col_Inv_DueDate.Index).Value = Format(.Inv_Due_Date, "dd MMM yyyy")
                        rowDGV.Cells(Frm.ColCustCode.Index).Value = objCust.GetCustomerCode(.Primary_Cust)
                        rowDGV.Cells(Frm.ColCustName.Index).Value = objCust.GetCustName(.Primary_Cust)
                        rowDGV.Cells(Frm.ColSales.Index).Value = ObjCD.DigitGrouping(.Sales)
                        rowDGV.Cells(Frm.ColVAT.Index).Value = ObjCD.DigitGrouping(.VAT)
                        rowDGV.Cells(Frm.ColDutystamp.Index).Value = ObjCD.DigitGrouping(.Dutystamp)
                        rowDGV.Cells(Frm.ColReceived.Index).Value = ObjCD.DigitGrouping(.Received)
                        rowDGV.Cells(Frm.ColBalance.Index).Value = ObjCD.DigitGrouping(.Balance)
                        rowDGV.Cells(Frm.ColTotal.Index).Value = ObjCD.DigitGrouping(.Total)
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
    ByVal Frm As FrmTransaksiSalesBalance_Import) As Boolean
        Dim row As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseTransaksiSalesBalance_ImportTableAdapters.tbltrans_salesbalance_import_itemTableAdapter
        Dim DefaultPrecision As Integer = 0
        Dim DueDate As Date = Today
        Try
            'simpan data Produk Code No
            For Each row In vDGVItem.Rows
                If String.IsNullOrEmpty(row.Cells(Frm.ColRemarks.Index).Value) = True Then
                    row.Cells(Frm.ColRemarks.Index).Value = ""
                End If

                If row.Cells(Frm.ColInv_Number.Index).Value <> "" Then
                    'simpan ke basis data item
                    DTA.Insert(vPrimaryTransaksi, CDbl(row.Cells(Frm.ColPrimaryInv.Index).Value), row.Cells(Frm.ColInv_Number.Index).Value, _
                     CDate(row.Cells(Frm.ColInv_Date.Index).Value), CDate(row.Cells(Frm.Col_Inv_DueDate.Index).Value), CDbl(row.Cells(Frm.ColPrimaryCust.Index).Value), _
                     CDbl(row.Cells(Frm.ColSales.Index).Value), CDbl(row.Cells(Frm.ColVAT.Index).Value), CDbl(row.Cells(Frm.ColDutystamp.Index).Value), _
                     CDbl(row.Cells(Frm.ColReceived.Index).Value), CDbl(row.Cells(Frm.ColBalance.Index).Value), row.Cells(Frm.ColRemarks.Index).Value, _
                     CDbl(row.Cells(Frm.ColTotal.Index).Value))

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
        Dim DTA As New DseTransaksiSalesBalance_ImportTableAdapters.tbltrans_salesbalance_import_itemTableAdapter
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


End Class
