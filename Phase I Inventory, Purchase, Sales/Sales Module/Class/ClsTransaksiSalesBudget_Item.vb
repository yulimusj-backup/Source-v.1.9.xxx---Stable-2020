Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesBudget_Item
    Public Function IsiDGV(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vFrm As FrmTransaksiSalesBudget) As Boolean
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.DataTable1TableAdapter
        Dim TBL As New DseTransaksiSalesBudget.DataTable1DataTable
        Dim row As DseTransaksiSalesBudget.DataTable1Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DefPrec As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DefPrec = ObjCur.GetDecimal_DefaultCurrency

            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    vDGVItem.Rows.Add()
                    rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                    With rowDGV
                        .Cells(vFrm.ColJan.Index).Value = ObjCD.DigitGrouping(row.jan, DefPrec)
                        .Cells(vFrm.ColFeb.Index).Value = ObjCD.DigitGrouping(row.feb, DefPrec)
                        .Cells(vFrm.ColMar.Index).Value = ObjCD.DigitGrouping(row.mar, DefPrec)
                        .Cells(vFrm.ColApr.Index).Value = ObjCD.DigitGrouping(row.apr, DefPrec)
                        .Cells(vFrm.ColMay.Index).Value = ObjCD.DigitGrouping(row.mei, DefPrec)
                        .Cells(vFrm.ColJun.Index).Value = ObjCD.DigitGrouping(row.jun, DefPrec)
                        .Cells(vFrm.ColJul.Index).Value = ObjCD.DigitGrouping(row.jul, DefPrec)
                        .Cells(vFrm.ColAug.Index).Value = ObjCD.DigitGrouping(row.agt, DefPrec)
                        .Cells(vFrm.ColSep.Index).Value = ObjCD.DigitGrouping(row.sep, DefPrec)
                        .Cells(vFrm.ColOct.Index).Value = ObjCD.DigitGrouping(row.okt, DefPrec)
                        .Cells(vFrm.ColNov.Index).Value = ObjCD.DigitGrouping(row.nov, DefPrec)
                        .Cells(vFrm.ColDec.Index).Value = ObjCD.DigitGrouping(row.des, DefPrec)

                        .Cells(vFrm.ColNo.Index).Value = vDGVItem.Rows.Count
                        .Cells(vFrm.ColPrimarySalesPerson.Index).Value = row.primary_sales_person
                        .Cells(vFrm.ColRemarks.Index).Value = row.remarks
                        .Cells(vFrm.ColSalesPersonCode.Index).Value = row.Sales_Person_Code
                        .Cells(vFrm.ColSalesPersonName.Index).Value = row.Sales_Person_Desc
                        .Cells(vFrm.ColTotal.Index).Value = ObjCD.DigitGrouping(row.jan + row.feb + row.mar + row.apr + row.mei + row.jun + _
                        row.jul + row.agt + row.sep + row.okt + row.nov + row.des, DefPrec)
                        .Cells(vFrm.ColSalesPGroup.Index).Value = row.Sales_Group_Name
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
            ObjCur = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudget_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vFrm As FrmTransaksiSalesBudget) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiSalesBudgetTableAdapters.tbltrans_salesbudget_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each row In vdgvItem.Rows
                With row
                    DTA.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColPrimarySalesPerson.Index).Value, CDec(row.Cells(vFrm.ColJan.Index).Value), _
                    CDec(row.Cells(vFrm.ColFeb.Index).Value), CDec(row.Cells(vFrm.ColMar.Index).Value), CDec(row.Cells(vFrm.ColApr.Index).Value), _
                    CDec(row.Cells(vFrm.ColMay.Index).Value), CDec(row.Cells(vFrm.ColJun.Index).Value), CDec(row.Cells(vFrm.ColJul.Index).Value), _
                    CDec(row.Cells(vFrm.ColAug.Index).Value), CDec(row.Cells(vFrm.ColSep.Index).Value), CDec(row.Cells(vFrm.ColOct.Index).Value), _
                    CDec(row.Cells(vFrm.ColNov.Index).Value), CDec(row.Cells(vFrm.ColDec.Index).Value), row.Cells(vFrm.ColRemarks.Index).Value)
                End With
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
