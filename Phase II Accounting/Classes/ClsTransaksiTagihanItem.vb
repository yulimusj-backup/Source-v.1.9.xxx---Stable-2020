Option Strict Off
Option Explicit On

Public Class ClsTransaksiTagihanItem
    Public Function IsiDGVItem(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTransaksiTagihan) As Boolean
        Dim DTA As New DseTransaksiTagihanTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiTagihan.DataTable2DataTable
        Dim row As DseTransaksiTagihan.DataTable2Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim obj As New ClsTransaksiTagihan
        Dim DefaultPrecision As Integer = 0
        Try
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()

            'isikan tabel
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add()
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColNo.Index).Value = ""
                        rowDGV.Cells(Frm.ColInvoiceNumber.Index).Value = .Invoice_Number
                        rowDGV.Cells(Frm.ColInvoiceDate.Index).Value = Format(.Invoice_Date, "dd MMM yyyy")
                        rowDGV.Cells(Frm.ColDueDate.Index).Value = Format(.DueDate, "dd MMM yyyy")
                        rowDGV.Cells(Frm.ColCurrency.Index).Value = .Currency_Code
                        rowDGV.Cells(Frm.ColAmount.Index).Value = ObjCD.DigitGrouping(.Total)
                        rowDGV.Cells(Frm.ColRemarks.Index).Value = .Remarks
                        rowDGV.Cells(Frm.ColPrimaryInvoice.Index).Value = .Primary_Invoice
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
            obj = Nothing
        End Try
    End Function

    Public Function SimpanItem(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmTransaksiTagihan) As Boolean
        Dim row As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihan_itemTableAdapter
        Dim DefaultPrecision As Integer = 0
        Try
            'simpan data Produk Code No
            For Each row In vDGVItem.Rows
                If String.IsNullOrEmpty(row.Cells(Frm.ColRemarks.Index).Value) = True Then
                    row.Cells(Frm.ColRemarks.Index).Value = ""
                End If
                If row.Cells(Frm.ColInvoiceNumber.Index).Value <> "" Then
                    'simpan ke basis data item
                    DTA.Insert(vPrimaryTransaksi, CDbl(row.Cells(Frm.ColPrimaryInvoice.Index).Value), _
                    row.Cells(Frm.ColRemarks.Index).Value)
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

    Public Function hapusitem(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim dta As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihan_itemTableAdapter
        Try
            'hapus data item
            dta.Delete(vPrimaryTransaksi)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function


End Class
