Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesInvoice
    Public Enum EnumJenisFakturPajak As Integer
        FakturPajakSederhana
        FakturPajakStandar
    End Enum
    Public Enum EnumJenisFakturStandar
        HargaJual
        Penggantian
        UangMuka
        Termin
    End Enum

    Public Function CetakInvoiceDPCustom(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vprimarysalesInvoice As Double, _
    ByVal vTanggalInvoice As Date, ByVal DecimalHarga As Integer) As Boolean

        Return True
    End Function

    Public Function CetakCustomVAT_DP(ByVal vDecimalValas As Double, ByVal vNoSeriFakturPajak As String, ByVal vNoFakturPajak As String, _
    ByVal vNoInvoice As String, ByVal vCustName As String, ByVal vPrimaryInvoice As Double, ByVal vNoNPWP As String, _
    ByVal vTotalHargaJual As Decimal, ByVal vTaxAmount As Decimal, ByVal vTaxExchangeRate As Decimal, ByVal vCurCode As String, _
    ByVal vDate As Date, ByVal vSONumber As String, ByVal vPrimaryCurrency As Double, ByVal vTax As Decimal, _
    ByVal vPrimarySO As Double, ByVal vFaktukPajakStandarValas As Boolean, ByVal vNamaSigned As String, _
    ByVal vJabatanSigned As String, ByVal vRatePajak As Decimal, ByVal vSubTotal As Decimal, ByVal vPersenPajak As Decimal, _
    ByVal vFormatDay As Integer, ByVal vFormatMonth As Integer, ByVal vFormatYears As Integer, ByVal vCurrencyCodeValas As String, _
    ByVal vFormatQty As Integer, ByVal vCrvFaktur As CrystalDecisions.Windows.Forms.CrystalReportViewer) As Boolean

    End Function

    Public Sub CetakCustomVAT(ByVal DecimalValas As Double, ByVal vPrimaryInv As Double, _
    ByVal vCurrencyCode As String, ByVal FakturPajakStandarValas As Boolean, ByVal RatePajak As Decimal, _
    ByVal PersenPajak As Decimal, ByVal JenisFakturPajakStandar As EnumJenisFakturStandar, ByVal FormatDay As Integer, _
    ByVal FormatMonth As Integer, ByVal FormatYears As Integer, ByVal FormatQty As Integer, ByVal CurrencyCodeValas As String, _
    ByVal FormatDecimal As Integer, ByVal TanggalStr As String, ByVal ObjCD As Core_Module.ClsCekData, _
    ByVal CRVfaktur As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal IsShipmentWithTax As Boolean)

    End Sub

    Public Function CetakInvoiceCustom(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vprimarysalesInvoice As Double, ByVal vTanggalInvoice As Date, ByVal DecimalHarga As Integer, ByVal vParameter_FormatDay As String, _
    ByVal vParameter_FormatMonth As String, _
    ByVal vParameter_FormatYears As String, _
    ByVal vParameter_FormatDecimal As String, _
    ByVal vParameter_Tanggal_Invoice As String, _
    ByVal vParameter_Tanggal_STR As String, _
    ByVal vParameter_Jabatan As String, _
    ByVal vParameter_Remarks As String) As Boolean
        Dim DS As New DseTransaksiSalesInvoice
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim ListStr As New List(Of String)
        Dim rowIsi As DseTransaksiSalesInvoice.TblCetakIsiRow
        Dim rowHeader As DseTransaksiSalesInvoice.TblCetakHeaderRow
        Dim counter As Integer = 0
        Dim Dummy As String = ""
        Dim DTAReturn As New DseTransaksiSalesInvoiceTableAdapters.DataTable9TableAdapter
        Dim TBLReturn As New DseTransaksiSalesInvoice.DataTable9DataTable
        Dim rowReturn As DseTransaksiSalesInvoice.DataTable9Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim report As New CrpCetakTransaksiInvoice
        Dim ObjCustom As New CustomSMESS.ClsTransaksiSalesInvoice
        Dim OngkosKirim As Decimal = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            Dim DTAHeader As New DseTransaksiSalesInvoiceTableAdapters.TblCetakHeaderTableAdapter
            DTAHeader.Connection.ConnectionString = Core_Module.GetConstring
            DTAHeader.FillBy_PrimaryInvoice(DS.TblCetakHeader, vprimarysalesInvoice)
            DTAHeader.Dispose()
            Dim DTAIsi As New DseTransaksiSalesInvoiceTableAdapters.TblCetakIsiTableAdapter
            DTAIsi.Connection.ConnectionString = Core_Module.GetConstring
            DTAIsi.Fill(DS.TblCetakIsi, vprimarysalesInvoice)
            DTAIsi.Dispose()

            'kurangkan isi yang terjadi dengan sales return
            DTAReturn.Connection.ConnectionString = Core_Module.GetConstring
            If DTAReturn.FillBy_PrimaryInvoice(TBLReturn, vprimarysalesInvoice) > 0 Then
                Dim DTAReturn_Detail As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_item_view0TableAdapter
                DTAReturn_Detail.Connection.ConnectionString = Core_Module.GetConstring
                Dim TBLReturn_Detail As New dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0DataTable
                Dim rowReturnDetail As dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0Row

                For Each rowReturn In TBLReturn.Rows
                    DTAReturn_Detail.Fill(TBLReturn_Detail, rowReturn.Primary_SR)
                    For Each rowReturnDetail In TBLReturn_Detail.Rows
                        Me.KurangkanItemInvoice(rowReturnDetail, DS.TblCetakIsi, DecimalHarga)
                    Next
                Next

                DTAReturn_Detail.Dispose()
                TBLReturn_Detail.Dispose()
            End If

            'tambahkan ongkos kirim
            OngkosKirim = Me.GetOngkosKirim(vprimarysalesInvoice)
            If OngkosKirim <> 0 And Me.IsShipmentWithTax(vprimarysalesInvoice) = True Then
                rowIsi = DS.TblCetakIsi.NewTblCetakIsiRow
                With rowIsi
                    rowIsi.Amount = OngkosKirim
                    .Category_Code = ""
                    .Category_Desc = ""
                    .Cust_PO_No = ""
                    .Description = "Ongkos Kirim"
                    .Discount = 0
                    .Nett_Price = OngkosKirim
                    .No_DO = ""
                    .No_SO = ""
                    .PaymentTermSO = ""
                    .Primary_Category = -1
                    .Primary_DO = -1
                    .Primary_Product = -1
                    .Primary_SO = -1
                    .Primary_Transaksi = -1
                    .Product_Code = "OK"
                    .Qty_36 = 0
                    .Qty_37 = 0
                    .Qty_38 = 0
                    .Qty_39 = 0
                    .Qty_40 = 0
                    .Qty_41 = 0
                    .Qty_42 = 0
                    .Qty_43 = 0
                    .Qty_44 = 0
                    .Qty_45 = 0
                    .Qty_46 = 0
                    .Remarks = ""
                    .Sec_Unit = ""
                    .Secondary_Qty = -1
                    .Size = 0
                    .TglSO = Today
                    .Total_Qty = 1
                    .Unit_Code = "UNIT"
                    .Unit_Price = OngkosKirim
                    .Use_Sizing = 0
                    ._Decimal = 0
                    .Alias_Code = ""
                    .Alias_Desc = ""
                End With
                DS.TblCetakIsi.Rows.Add(rowIsi)
            End If

            'masukkan kode alias
            For Each rowIsi In DS.TblCetakIsi.Rows
                If rowIsi.IsAlias_CodeNull = False Then
                    If rowIsi.Alias_Code.Trim <> "" Then
                        rowIsi.Product_Code = rowIsi.Alias_Code.Trim
                    End If
                End If
                If rowIsi.IsAlias_DescNull = False Then
                    If rowIsi.Alias_Desc.Trim <> "" Then
                        rowIsi.Description = rowIsi.Alias_Desc.Trim
                    End If
                End If
            Next

            'update data customer order
            For Each rowIsi In DS.TblCetakIsi.Rows
                If ListStr.Contains(rowIsi.No_DO) = False And rowIsi.No_DO <> "" Then
                    ListStr.Add(rowIsi.No_DO)
                End If
            Next
            rowHeader = DS.TblCetakHeader.Rows(0)
            If rowHeader.Shipment_With_VAT = 1 Then
                rowHeader.ongkos_kirim = 0
                rowHeader.Sub_Total = rowHeader.Sub_Total + OngkosKirim
            End If
            If rowHeader.Currency_Code = "IDR" Then
                rowHeader.Currency_Code = ""
            Else
                If rowHeader.VAT_Dipungut = 1 Then
                    rowHeader.Tax_Amount = Math.Round((rowHeader.Sub_Total - rowHeader.Down_Payment) * rowHeader.Tax / 100, ObjCur.GetDefaultPrecission(rowHeader.Primary_Currency), MidpointRounding.AwayFromZero)
                End If
                If rowHeader.Shipment_With_VAT = 1 Then
                    rowHeader.Total = rowHeader.Sub_Total + rowHeader.Tax_Amount - rowHeader.Down_Payment
                Else
                    rowHeader.Total = rowHeader.Sub_Total + rowHeader.Tax_Amount + OngkosKirim - rowHeader.Down_Payment
                End If
            End If

            'modifikasi item sehingga untuk product yang sama tetapi batch, unit price, berbeda tidak keluar 2 kali.
            Me.PindahkanItem(DS)

            For Each Dummy In ListStr
                counter += 1
                If rowHeader.IsNo_DO_Buatan_SendiriNull = True Then
                    If counter = ListStr.Count Then
                        rowHeader.No_DO_Buatan_Sendiri = Dummy
                    Else
                        rowHeader.No_DO_Buatan_Sendiri = Dummy.Substring(0, 6)
                    End If
                Else
                    If counter = ListStr.Count Then
                        rowHeader.No_DO_Buatan_Sendiri = rowHeader.No_DO_Buatan_Sendiri + ", " + Dummy
                    Else
                        rowHeader.No_DO_Buatan_Sendiri = rowHeader.No_DO_Buatan_Sendiri + ", " + Dummy.Substring(0, 6)
                    End If
                End If
            Next

            'cust po number
            ListStr.Clear()
            For Each rowIsi In DS.TblCetakIsi.Rows
                If ListStr.Contains(rowIsi.Cust_PO_No) = False And rowIsi.Cust_PO_No <> "" Then
                    ListStr.Add(rowIsi.Cust_PO_No)
                End If
            Next
            counter = 0
            For Each Dummy In ListStr
                counter += 1
                If rowHeader.IsNo_PO_Buatan_SendiriNull = True Then
                    If counter = ListStr.Count Then
                        rowHeader.No_PO_Buatan_Sendiri = Dummy
                    Else
                        rowHeader.No_PO_Buatan_Sendiri = Dummy
                    End If
                Else
                    If counter = ListStr.Count Then
                        rowHeader.No_PO_Buatan_Sendiri = rowHeader.No_PO_Buatan_Sendiri + ", " + Dummy
                    Else
                        rowHeader.No_PO_Buatan_Sendiri = rowHeader.No_PO_Buatan_Sendiri + ", " + Dummy
                    End If
                End If
            Next

            If rowHeader.IsNo_PO_Buatan_SendiriNull = True Then
                rowHeader.No_PO_Buatan_Sendiri = ""
            End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            report.SetParameterValue(report.Parameter_FormatDecimalIDR.ParameterFieldName, ObjCur.GetDecimal_DefaultCurrency)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, ObjCur.GetDefaultPrecission(rowHeader.Primary_Currency))
            report.SetParameterValue(report.Parameter_Tanggal_Invoice.ParameterFieldName, vTanggalInvoice)
            report.SetParameterValue(report.Parameter_Tanggal_STR.ParameterFieldName, ObjCD.GetDateString_Indo(vTanggalInvoice))
            report.SetParameterValue(report.Parameter_Jabatan.ParameterFieldName, rowHeader.Jabatan_Signed)
            report.SetParameterValue(report.Parameter_Remarks.ParameterFieldName, rowHeader.Remarks)

            CRV.ReportSource = report
            CRV.Show()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTAReturn.Dispose()
            TBLReturn.Dispose()
            ObjCD = Nothing
            ObjSetting = Nothing
            ObjCustom = Nothing
            ObjCur = Nothing
        End Try
    End Function

    Private Sub PindahkanItem(ByVal DS As DseTransaksiSalesInvoice)
        Dim TBLIsi_Dummy As New DseTransaksiSalesInvoice.TblCetakIsiDataTable
        Dim rowIsi_Dummy, rowIsi As DseTransaksiSalesInvoice.TblCetakIsiRow
        Dim Ketemu As Boolean = False

        Try
            For Each rowIsi In DS.TblCetakIsi.Rows
                rowIsi_Dummy = TBLIsi_Dummy.NewTblCetakIsiRow
                With rowIsi_Dummy
                    .Amount = rowIsi.Amount
                    .Cust_PO_No = rowIsi.Cust_PO_No
                    .Description = rowIsi.Description
                    .Discount = rowIsi.Discount
                    .Nett_Price = rowIsi.Nett_Price
                    .No_DO = rowIsi.No_DO
                    .No_SO = rowIsi.No_SO
                    .PaymentTermSO = rowIsi.PaymentTermSO
                    .Primary_Product = rowIsi.Primary_Product
                    .Primary_SO = rowIsi.Primary_SO
                    .Primary_Transaksi = rowIsi.Primary_Transaksi
                    .Product_Code = rowIsi.Product_Code
                    .Remarks = rowIsi.Remarks
                    .Sec_Unit = rowIsi.Sec_Unit
                    .Secondary_Qty = rowIsi.Secondary_Qty
                    .TglSO = rowIsi.TglSO
                    .Total_Qty = rowIsi.Total_Qty
                    .Unit_Code = rowIsi.Unit_Code
                    .Unit_Price = rowIsi.Unit_Price
                    ._Decimal = rowIsi._Decimal
                    .Use_Sizing = rowIsi.Use_Sizing
                    .Size = rowIsi.Size
                    .Category_Code = rowIsi.Category_Code
                    .Category_Desc = rowIsi.Category_Desc
                    .Primary_DO = rowIsi.Primary_DO
                    .Primary_Category = rowIsi.Primary_Category
                End With
                TBLIsi_Dummy.Rows.Add(rowIsi_Dummy)
            Next
            DS.TblCetakIsi.Rows.Clear()
            For Each rowIsi In TBLIsi_Dummy.Rows
                Ketemu = False
                For Each rowIsi_Dummy In DS.TblCetakIsi.Rows
                    If rowIsi_Dummy.No_SO = rowIsi.No_SO And rowIsi_Dummy.No_DO = rowIsi.No_DO _
                    And rowIsi_Dummy.Primary_Product = rowIsi.Primary_Product _
                    And rowIsi_Dummy.Unit_Price = rowIsi.Unit_Price _
                    And rowIsi_Dummy.Discount = rowIsi.Discount _
                    And rowIsi_Dummy.Nett_Price = rowIsi.Nett_Price Then
                        'tandai sudah ketemu
                        Ketemu = True

                        'modifikasi qty dan harga
                        With rowIsi_Dummy
                            .Total_Qty += rowIsi.Total_Qty
                            .Amount = .Amount + (rowIsi.Amount)
                        End With

                        Exit For
                    End If
                Next

                If Ketemu = False Then
                    rowIsi_Dummy = DS.TblCetakIsi.NewTblCetakIsiRow
                    With rowIsi_Dummy
                        .Amount = rowIsi.Amount
                        .Cust_PO_No = rowIsi.Cust_PO_No
                        .Description = rowIsi.Description
                        .Discount = rowIsi.Discount
                        .Nett_Price = rowIsi.Nett_Price
                        .No_DO = rowIsi.No_DO
                        .No_SO = rowIsi.No_SO
                        .PaymentTermSO = rowIsi.PaymentTermSO
                        .Primary_Product = rowIsi.Primary_Product
                        .Primary_SO = rowIsi.Primary_SO
                        .Primary_Transaksi = rowIsi.Primary_Transaksi
                        .Product_Code = rowIsi.Product_Code
                        .Remarks = rowIsi.Remarks
                        .Sec_Unit = rowIsi.Sec_Unit
                        .Secondary_Qty = rowIsi.Secondary_Qty
                        .TglSO = rowIsi.TglSO
                        .Total_Qty = rowIsi.Total_Qty
                        .Unit_Code = rowIsi.Unit_Code
                        .Unit_Price = rowIsi.Unit_Price
                        ._Decimal = rowIsi._Decimal
                        .Use_Sizing = rowIsi.Use_Sizing
                        .Size = rowIsi.Size
                        .Category_Code = rowIsi.Category_Code
                        .Category_Desc = rowIsi.Category_Desc
                        .Primary_DO = rowIsi.Primary_DO
                        .Primary_Category = rowIsi.Primary_Category
                    End With
                    DS.TblCetakIsi.Rows.Add(rowIsi_Dummy)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLIsi_Dummy.Dispose()
        End Try
    End Sub

    Public Function IsShipmentWithTax(ByVal vPrimaryInvoice As Double) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.TblCetakHeaderTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQueryGetShipmenWithVAT(vPrimaryInvoice).HasValue = True Then
                If DTA.ScalarQueryGetShipmenWithVAT(vPrimaryInvoice).Value = 1 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetOngkosKirim(ByVal vPrimaryInv As Double) As Decimal
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.TblCetakHeaderTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetOngkosKirim(vPrimaryInv).HasValue = True Then
                Return DTA.ScalarQuery_GetOngkosKirim(vPrimaryInv).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub KurangkanItemInvoice(ByVal rowReturn As dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0Row, _
  ByVal vTBLIsi As DseTransaksiSalesInvoice.TblCetakIsiDataTable, ByVal DefaultPrec As Integer)
        Dim rowInv As DseTransaksiSalesInvoice.TblCetakIsiRow
        Dim counter As Integer = 0
        Dim HargaLama, HargaBaru As Decimal

        Try
            While rowReturn.Total_Qty > 0
                While counter < vTBLIsi.Rows.Count
                    rowInv = vTBLIsi.Rows(counter)
                    If rowInv.Product_Code = rowReturn.Product_Code And rowInv.No_SO = rowReturn.SO_No And _
                    rowInv.No_DO = rowReturn.DO_Number Then
                        If rowInv.Total_Qty = rowReturn.Total_Qty Then
                            vTBLIsi.Rows.Remove(rowInv)
                            Exit Sub
                        ElseIf rowInv.Total_Qty < rowReturn.Total_Qty Then
                            rowReturn.Total_Qty -= rowInv.Total_Qty
                            vTBLIsi.Rows.Remove(rowInv)
                        ElseIf rowInv.Total_Qty > rowReturn.Total_Qty Then
                            rowInv.Total_Qty -= rowReturn.Total_Qty
                            'hitung ulang amount
                            HargaLama = rowInv.Amount
                            HargaBaru = Math.Round(rowInv.Amount * rowInv.Total_Qty / _
                            (rowInv.Total_Qty + rowReturn.Total_Qty), DefaultPrec, MidpointRounding.AwayFromZero)
                            rowInv.Amount = HargaBaru
                            Exit Sub
                        End If
                    Else
                        counter += 1
                    End If
                End While
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahBolehHapus_Custom(ByVal vPrimaryInvoiceSMESS As Double) As Boolean
        Return True
    End Function

    Public Function ApakahBolehRevisiCustom(ByVal vPrimaryInvoiceSMESS As Double) As Boolean
        Return True
    End Function

    Public Function ApakahBolehEdit_Custom(ByVal vPrimaryInvoiceSMESS As Double) As Boolean
        Return True
    End Function
End Class
