Option Explicit On
Option Strict Off

Public Class clsTransaksiSettingSupplierProductCodeNoItem
    Public Function IsiDGV(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiSettingSupplierProductCodeNoTableAdapters.tbltrans_supplierproductcodestting_item1TableAdapter
        Dim TBL As New DseTransaksiSettingSupplierProductCodeNo.tbltrans_supplierproductcodestting_item1DataTable
        Dim row As DseTransaksiSettingSupplierProductCodeNo.tbltrans_supplierproductcodestting_item1Row
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    vDGVItem.Rows.Add("0", row.Product_Code, row.Description, _
                    row.Specification, row._Supplier_s_Code_No, _
                    row._Supplier_s_Desc, row._Supplier_s_Specification, row.Remarks, row.Primary_Product)
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

    Public Function Simpan(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiSettingSupplierProductCodeNoTableAdapters.tbltrans_supplierproductcodestting_itemTableAdapter
        Try
            For Each row In vdgvItem.Rows
                DTA.Insert(vPrimaryTransaksi, row.Cells("ColPrimaryCodeNo").Value, row.Cells("ColSupplierCodeNo").Value, row.Cells("ColSuppDesc").Value, row.Cells("ColSuppSpecification").Value, row.Cells("ColRemark").Value)
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiSettingSupplierProductCodeNoTableAdapters.tbltrans_supplierproductcodestting_itemTableAdapter
        Try
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
