Option Explicit On
Option Strict Off

Public Class ClsTransaksiSOCItem
    Public Sub IsiTBL(ByVal vPrimarySOC As Double, ByVal TBL As DseTransaksiSOC.DataTable3DataTable)
        Dim DTA As New DseTransaksiSOCTableAdapters.DataTable3TableAdapter

        Try
            DTA.Fill(TBL, vPrimarySOC)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetJml_Default(ByVal vPrimarySOC As Double, ByVal vPrimarySO As Double, _
    ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New DseTransaksiSOCTableAdapters.tbltrans_soc_itemTableAdapter

        Try
            Return DTA.ScalarQuery_GetJml_Default(vPrimaryProduct, vPrimarySO, vPrimarySOC).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_Multi(ByVal vPrimarySOC As Double, ByVal vPrimarySO As Double, _
    ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New DseTransaksiSOCTableAdapters.tbltrans_soc_itemTableAdapter

        Try
            Return DTA.ScalarQuery_GetJml_Multi(vPrimaryProduct, vPrimarySO, vPrimarySOC).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV_Default(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiSOCTableAdapters.DataTable1TableAdapter
        Dim TBL As New DseTransaksiSOC.DataTable1DataTable
        Dim row As DseTransaksiSOC.DataTable1Row
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.FillBy_PrimarySOC(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("0", row.Product_Code, row.Description, _
                        ObjCD.DigitGrouping(.Qty, ._Decimal), row.Unit_Code, row.Remarks, row.Primary_Product, row.Primary_Units)
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
        End Try
    End Function

    Public Function IsiDGV_Multi(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiSOCTableAdapters.DataTable1TableAdapter
        Dim TBL As New DseTransaksiSOC.DataTable1DataTable
        Dim row As DseTransaksiSOC.DataTable1Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim RowDgv As DataGridViewRow
        Dim i As New FrmTransaksiSOC
        Dim DefPrecDefault As Integer = 0
        Dim DefPrecMulti As Integer = 0
        Dim ObjProduct As New Core_Module.clsMasterProduct

        Try
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.FillBy_PrimarySOC(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        DefPrecDefault = ._Decimal
                        DefPrecMulti = ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)
                        vDGVItem.Rows.Add()
                        RowDgv = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        RowDgv.Cells(i.ColNo.Index).Value = ""
                        RowDgv.Cells(i.ColCodeNo.Index).Value = .Product_Code
                        RowDgv.Cells(i.ColDescription.Index).Value = .Description
                        RowDgv.Cells(i.ColQty.Index).Value = ObjCD.DigitGrouping(.Qty, DefPrecDefault)
                        RowDgv.Cells(i.ColUnit.Index).Value = .Unit_Code
                        RowDgv.Cells(i.ColRemark.Index).Value = .Remarks
                        RowDgv.Cells(i.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        RowDgv.Cells(i.ColPrimaryUnits.Index).Value = .Primary_Units
                        RowDgv.Cells(i.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(.Qty_Multi, DefPrecMulti)
                        RowDgv.Cells(i.ColUnitsMulti.Index).Value = .Units_Code_Multi
                        RowDgv.Cells(i.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(.Ratio_Default, DefPrecDefault)
                        RowDgv.Cells(i.ColRatioReference.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti)
                        If row.Primary_Units <> row.Primary_Units_Multi Then
                            RowDgv.Cells(i.ColRatio.Index).Value = ObjCD.DigitGrouping(row.Ratio_Reference, _
                                    ObjProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Units_Multi)) _
                                    & " " & row.Units_Code_Multi & " = " & ObjCD.DigitGrouping(row.Ratio_Default, _
                                    row._Decimal) & " " & row.Unit_Code
                        Else
                            RowDgv.Cells(i.ColRatio.Index).Value = "-"
                        End If
                        RowDgv.Cells(i.ColPrimaryUnitsMulti.Index).Value = .Primary_Units_Multi
                        'vDGVItem.Rows.Add("0", row.Product_Code, row.Description, _
                        'ObjCD.DigitGrouping(.Qty, ._Decimal), row.Unit_Code, row.Remarks, row.Primary_Product, row.Primary_Units)
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
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiSOCTableAdapters.tbltrans_soc_itemTableAdapter
        Dim i As New FrmTransaksiSOC

        Try
            For Each row In vdgvItem.Rows
                'simpan ke basis data item
                DTA.Insert(vPrimaryTransaksi, row.Cells(i.ColPrimaryCodeNo.Index).Value, row.Cells(i.ColQty.Index).Value, row.Cells(i.ColPrimaryUnits.Index).Value, _
                row.Cells(i.ColRemark.Index).Value, row.Cells(i.ColQtyMulti.Index).Value, row.Cells(i.ColPrimaryUnitsMulti.Index).Value, row.Cells(i.ColRatioReference.Index).Value, _
                row.Cells(i.ColRatioDefault.Index).Value)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiSOCTableAdapters.tbltrans_soc_itemTableAdapter

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
