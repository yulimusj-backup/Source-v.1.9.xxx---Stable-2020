Option Explicit On
Option Strict Off

Public Class FrmTransaksiStockInSalesReturn2_Print
    Public primary As Double = -1
    'Public TipeLaporan As EnumTipeLaporan = EnumTipeLaporan.StockInPurchase
    'Public FormatLaporan As EnumFormatLaporan = EnumFormatLaporan.A4
    'Dim ObjCd As New Core_Module.ClsCekData
    'Public UserCode As String = ""
    'Public SudahAdaStockOut As Boolean = False
    'Dim Print As Boolean = False
    'Dim ObjSetting As New Core_Module.ClsTransaksiSetting

    'Dim DTADO As New DseTransaksiDOTableAdapters.tbltrans_do5TableAdapter
    'Dim TBLDO As New DseTransaksiDO.tbltrans_do5DataTable
    'Dim rowDO As DseTransaksiDO.tbltrans_do5Row
    Dim objSIP As New CustomSMESS.ClsTransaksiStockInOut

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmCetakStockInPurchase_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmCetakStockInPurchase_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        'Dim report As New CrpCetakTransaksiStockInPurchase
        'Dim ds As New DseTransaksiStockInOut
        'Dim dta As New DseTransaksiStockInOutTableAdapters.DataTable1TableAdapter
        'Dim TBL As DseTransaksiStockInOut.DataTable1DataTable = ds.DataTable1
        'Dim row As DseTransaksiStockInOut.DataTable1Row
        'Dim dta2 As New DseTransaksiStockInOutTableAdapters.DataTable2TableAdapter
        'Dim rowBefore As DseTransaksiStockInOut.DataTable1Row
        'Windows.Forms.Cursor.Current = Cursors.WaitCursor
        'Dim ObjCD As New Core_Module.ClsCekData
        'Dim ObjProduct As New Core_Module.clsMasterProduct
        'Dim maxRow As Integer = 0
        'Dim limit As Integer = 0
        'Dim VirtualRow As Integer = 0
        'Dim lastProductCode As String = ""
        'Dim Value_2 As String = ""
        'Try
        '    dta.Connection.ConnectionString = Core_Module.GetConstring
        '    dta2.Connection.ConnectionString = Core_Module.GetConstring
        '    If dta.FillBy_PrimaryStockIn(TBL, primary) > 0 Then

        '        maxRow = 11
        '        limit = maxRow
        '        For Each row In TBL
        '            With row
        '                ._Decimal = ObjProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Multi_Units)
        '                ''filter berdasarkan primary product dan primary kompetensi
        '                If dta2.ScalarQuery_GetValueByPrimaryProductNKompetensi_PcsCtn(.Primary_Product) > 0 Then
        '                    Value_2 = dta2.ScalarQuery_GetValueByPrimaryProductNKompetensi_PcsCtn(.Primary_Product)
        '                    .Value_ = Value_2
        '                End If
        '                ''pagination
        '                VirtualRow += 1
        '                If TBL.Rows.IndexOf(row) > 0 Then
        '                    rowBefore = TBL.Rows(TBL.Rows.IndexOf(row) - 1)
        '                    If String.IsNullOrEmpty(row.Batch_No) = False Then
        '                        If rowBefore.Primary_Product <> .Primary_Product Then
        '                            VirtualRow += 2
        '                        End If
        '                    End If
        '                End If
        '                If VirtualRow > limit Then
        '                    limit += maxRow
        '                End If
        '                row.Group_ = CInt(limit / maxRow).ToString
        '                TBL.DefaultView.RowFilter = "[Group_] = " & CDec(.Group_).ToString & " and [Primary Product] = " & .Primary_Product.ToString
        '                If TBL.DefaultView.Count > 0 Then
        '                    row = TBL.DefaultView(0).Row
        '                    row.Qty_ += .Qty_Multi_Unit
        '                Else
        '                    row = TBL.NewDataTable1Row
        '                    row.Primary_Product = .Primary_Product
        '                    row.Qty_ = .Qty_Multi_Unit
        '                    row.Group_ = .Group_
        '                    TBL.Rows.Add(row)
        '                End If
        '                lastProductCode = .Product_Code
        '            End With
        '        Next
        '        'tambah baris kosong untuk mengisi sampai ujung bawah kertas
        '        If VirtualRow < limit Then
        '            For i As Integer = 1 To limit - VirtualRow
        '                row = TBL.NewDataTable1Row
        '                With row
        '                    .Primary = -1
        '                    .No_Transaksi = ""
        '                    .Tanggal_Transaksi = Now
        '                    ._Supplier_s_DO_Number = ""
        '                    .Supplier_Code = ""
        '                    .Supplier_Name = ""
        '                    .No_Po = ""
        '                    .Remarks = ""
        '                    .Primary_Product = -1
        '                    .Product_Code = lastProductCode + CStr(99 * i)
        '                    .Desc_Product = ""
        '                    .Batch_No = ""
        '                    .Expiry_Date = Now
        '                    .Primary_Multi_Units = -1
        '                    .Unit_Code = ""
        '                    .Qty_Multi_Unit = 0
        '                    ._Decimal = 0
        '                    .Value_ = ""
        '                    .Decimal_MultiUnit_ = 0
        '                    .Group_ = CInt(limit / maxRow).ToString
        '                    .Qty_ = 0
        '                End With
        '                TBL.Rows.Add(row)
        '            Next
        '        End If
        '        report.SetDataSource(ds)
        '        report.SetParameterValue(report.Parameter_Last_Group.ParameterFieldName, CInt(limit / maxRow).ToString)
        '        Me.CrvStockInPurchase.ReportSource = report
        '        Me.CrvStockInPurchase.Refresh()
        '        Me.CrvStockInPurchase.Show()
        '    End If

        '    Windows.Forms.Cursor.Current = Cursors.Arrow
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    dta.Dispose()
        '    TBL.Dispose()
        '    dta2.Dispose()
        '    ds.Dispose()
        '    ObjCD = Nothing
        'End Try
    End Sub
End Class