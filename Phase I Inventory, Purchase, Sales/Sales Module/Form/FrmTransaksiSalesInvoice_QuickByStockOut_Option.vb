Option Explicit On
Option Strict Off

Public Class FrmTransaksiSalesInvoice_QuickByStockOut_Option
    Public ListPrimaryStockOutSales As New List(Of Double)
    Public tblOutSales As New DseTransaksiSalesInvoice_QuickByStockOut.tbl_headerDataTable
    Public tblCreate As New DseTransaksiSalesInvoice_QuickByStockOut.tbl_create_invoiceDataTable
    Dim rowOutSales As DseTransaksiSalesInvoice_QuickByStockOut.tbl_headerRow
    Dim rowCreate As DseTransaksiSalesInvoice_QuickByStockOut.tbl_create_invoiceRow
    Dim ObjInv As New ClsTransaksiSalesInvoice
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCustomer As New Core_Module.clsMasterCustomer
    Dim ObjLicense As New ClsLicense
    Dim ObjCL As New ClsTransaksiPurchaseClosing
    Dim ObjTaxInvoiceSetting As New ClsTransaksiSalesInvoice_SettingTaxInvoice
    Dim ObjTaxRate As New ClsTransaksiSalesInvoiceTaxRate

    Private Sub FrmTransaksiSalesInvoice_QuickByStockOut_Option_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ChkUseCustomPrinter.Enabled = False
            Me.ChkDirectlyPrint.Enabled = False
            Me.SetCustomPrinter(False)

            Me.BtnStart.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkUseCustomPrinter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkUseCustomPrinter.CheckedChanged
        Try
            Me.SetCustomPrinter(Me.ChkUseCustomPrinter.Checked)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetCustomPrinter(ByVal vUseCustomPrinter As Boolean)
        Try
            Me.BtnFind.Enabled = vUseCustomPrinter
            If vUseCustomPrinter = False Then
                'Me.TxtUseCustomPrinter.Text  = SettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.BAGUS_PrinterName_QuickStockOutSales, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PrinterName, My.Computer.Name)
                Me.TxtUseCustomPrinter.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Me.Close()
    End Sub

    Private Sub BtnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFind.Click
        Try
            If Me.PrintDialog1.ShowDialog() = Windows.Forms.DialogResult.Cancel Then
                If Me.TxtUseCustomPrinter.Text = "" Then
                    Me.TxtUseCustomPrinter.Text = ""
                End If
            Else
                Me.TxtUseCustomPrinter.Text = Me.PrintDialog1.PrinterSettings.PrinterName
            End If
            'Me.PrintDialog1.ShowDialog()
            'Me.TxtUseCustomPrinter.Text = Me.PrintDialog1.PrinterSettings.PrinterName
            'SettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.BAGUS_PrinterName, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PrinterName, Me.PrintDialog1.PrinterSettings.PrinterName, My.Computer.Name)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnStart.Click
        Dim Tbl As New DseTransaksiSalesInvoice_QuickByStockOut.tbl_exceptionDataTable
        Dim ij As New FrmTransaksiSalesInvoice_QuickByStockOut_Exception
        Dim NoFakturEndingMax As String = ""
        Dim NoFakturDigunakanMax As String = ""
        Dim NoFakturTersedia As Integer = 0
        Dim TanggalClosingMax As Date = Now
        Dim ListPrimaryCustomer As New List(Of Double)
        Dim ListPrimaryCurr_OtherIDR As New List(Of Double)
        Dim Urutan As Integer = 0
        Dim frmInv As New FrmTransaksiSalesInvoice
        Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
        Dim hasil As MsgBoxResult

        Try
            'NoFakturEndingMax = Me.GetNoSeriFakturPajakEndingNumber(Me.DtpDate.Value.Date)
            'NoFakturDigunakanMax = Me.ObjInv.GetNoFakturPajakStandarBaru(Me.DtpDate.Value.Date)
            'If IsNumeric(NoFakturDigunakanMax) = False Then
            '    NoFakturDigunakanMax = 0
            'End If
            NoFakturTersedia = GetNoSeriFakturPajakTersedia(Me.DtpDate.Value.Date)

            If Me.ChkDirectlyPrint.Checked = True Or Me.ChkDirectlySave.Checked = True Then
                If ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = False Then
                    MsgBox("Sales invoice numbering is not automatic. You can't directly save or directly print.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            Urutan += 1
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.DtpDate.Value.Date Then
                Me.AddTblException(Tbl, -1, "", "This aplication has expire please contact your software vendor to renew your license.", Urutan)
            End If

            'cek berkaitan dengan tanggal closing
            Urutan += 1
            TanggalClosingMax = Me.ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod)
            If Me.DtpDate.Value.Date <= TanggalClosingMax.Date Then
                Me.AddTblException(Tbl, -1, "", "The period " + Format(Me.DtpDate.Value, "MMM yyyy") + " has been closed, you cannot save this data.", Urutan)
            End If

            Urutan += 1
            'tblCreate.DefaultView.RowFilter = "[use tax] = true"
            'If tblCreate.DefaultView.Count > NoFakturTersedia Then
            '    Me.AddTblException(Tbl, -1, "", "Sales invoice with tax will be generated is " & tblCreate.DefaultView.Count & ", while tax invoice no are available is " & NoFakturTersedia, Urutan)
            'End If
            'new aul
            tblCreate.DefaultView.RowFilter = "[use tax] = true"
            Dim ListInvoice As New List(Of Double)
            For Each dv As DataRowView In tblCreate.DefaultView
                rowCreate = dv.Row
                If ListInvoice.Contains(rowCreate.invoice_group) = False Then
                    ListInvoice.Add(rowCreate.invoice_group)
                End If
            Next
            If ListInvoice.Count > NoFakturTersedia Then
                Me.AddTblException(Tbl, -1, "", "Sales invoice with tax will be generated is " & tblCreate.DefaultView.Count & ", while tax invoice no are available is " & NoFakturTersedia, Urutan)
            End If

            Urutan += 1
            For Each rowOutSales In tblOutSales
                If ListPrimaryStockOutSales.Contains(rowOutSales.Primary) = True Then
                    If Format(rowOutSales.stockout_sales_date, "MMM yyyy") <> Format(DtpDate.Value.Date, "MMM yyyy") Then
                        Me.AddTblException(Tbl, rowOutSales.Primary, "", rowOutSales.stockout_sales_no & ": stock out sales must be dated on the same month as invoice date.", Urutan)
                    End If
                    If ListPrimaryCustomer.Contains(rowOutSales.primary_cust) = False Then
                        ListPrimaryCustomer.Add(rowOutSales.primary_cust)
                        Dim MsgException As String = ""
                        If rowOutSales.Tax_Add_Line1.Trim = "" And rowOutSales.Tax_Add_Line2.Trim = "" Then
                            If rowOutSales.use_tax = True Then
                                MsgException = rowOutSales.cust_code & "-" & rowOutSales.cust_name & ": please specify a valid tax address for the specified customer"
                            End If
                        End If
                        rowOutSales.NPWP_No = rowOutSales.NPWP_No.Trim
                        If Me.ObjCustomer.ApakahFormatNPWPBenar(rowOutSales.NPWP_No, False) = False Then
                            If MsgException = "" Then
                                MsgException &= rowOutSales.cust_code & "-" & rowOutSales.cust_name & ": please specify a valid NPWP format"
                            Else
                                MsgException &= " and please specify a valid NPWP format"
                            End If
                        End If
                        If MsgException <> "" Then
                            Me.AddTblException(Tbl, rowOutSales.Primary, "", MsgException, Urutan + 1)
                        End If
                    End If
                    If rowOutSales.primary_curr <> 1 Then
                        If ListPrimaryCurr_OtherIDR.Contains(rowOutSales.primary_curr) = False Then
                            ListPrimaryCurr_OtherIDR.Add(rowOutSales.primary_curr)
                        End If
                    End If
                End If
            Next

            If ListPrimaryCurr_OtherIDR.Count > 0 Then
                If Me.DtpDate.Value.Date >= Me.ObjTaxInvoiceSetting.BatasStartDate Then
                    If Me.ObjTaxRate.ApakahSudahAda(Me.DtpDate.Value.Date) = False Then
                        MsgBox("Please specify tax invoice rate data current transaction date.", MsgBoxStyle.Information, "Information")
                        frmInv.ShowTaxSetting()
                        Exit Sub
                    End If
                End If
            End If

            If Tbl.Rows.Count > 0 Then
                ij.Tbl = Tbl.Copy
                ij.ShowDialog()
                Exit Sub
            End If

            hasil = MsgBox("Create sales invoice?", MsgBoxStyle.YesNo, "Confirm")
            If hasil <> MsgBoxResult.Yes Then
                Exit Sub
            End If

            If Me.ChkUseCustomPrinter.Checked = True And Me.TxtUseCustomPrinter.Text <> "" Then
                'SettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.BAGUS_PrinterName_QuickStockOutSales, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PrinterName, Me.PrintDialog1.PrinterSettings.PrinterName, My.Computer.Name)
                'Me.PrinterName_ = SettingUser.GetSettingValue(s.ClsTransaksiSettingUser.EnumTipeTransaksi.BAGUS_PrinterName_QuickStockOutSales, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PrinterName, My.Computer.Name)
            End If

            Me.DialogResult = Windows.Forms.DialogResult.OK
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddTblException(ByVal Tbl As DseTransaksiSalesInvoice_QuickByStockOut.tbl_exceptionDataTable, _
    ByVal vPrimaryStockOutSales As Double, ByVal vNoOutSales As String, ByVal vReason As String, ByVal vUrutan As Integer)
        Dim row As DseTransaksiSalesInvoice_QuickByStockOut.tbl_exceptionRow

        Try
            row = Tbl.Newtbl_exceptionRow
            row.primary_stock_out_sales = vPrimaryStockOutSales
            row.stock_out_sales_no = vNoOutSales
            row.reason = vReason
            row.urutan = vUrutan
            Tbl.Rows.Add(row)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Public Function GetNoSeriFakturPajakEndingNumber(ByVal vTanggalTransaksi As Date) As Integer
    '    Dim DTA As New DseTransaksiSalesInvoice_QuickByStockOutTableAdapters.QueriesTableAdapter
    '    Dim NoUrut As Integer = 1
    '    Dim TglFrom As Date = Now
    '    Dim TglTo As Date = Now
    '    Dim ObjCD As New Core_Module.ClsCekData

    '    Try
    '        TglFrom = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)
    '        TglTo = TglFrom

    '        'cari no urut
    '        If DTA.ScalarQuery_GetMaxkUrutanInPeriod(1, TglFrom, TglTo).HasValue = True Then
    '            NoUrut = DTA.ScalarQuery_GetMaxkUrutanInPeriod(1, TglFrom, TglTo).Value + 1
    '        End If

    '        Return NoUrut
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return NoUrut
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Function GetNoSeriFakturPajakTersedia(ByVal vTanggalTransaksi As Date) As Integer
        Dim DTA As New DseTaxInvoiceSettingTableAdapters.tbltrans_salesinvoice_tax_invoice_settingTableAdapter
        Dim TBL As New DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingDataTable
        Dim row As DseTaxInvoiceSetting.tbltrans_salesinvoice_tax_invoice_settingRow
        Dim Number As Double = 0
        Dim ObjCD As New Core_Module.ClsCekData
        Dim Dummy As String = ""
        Dim StartingNumb As Integer = 0
        Dim NoFakturTersedia As Integer = 0

        Try
            'isikan berdasarkan tahun, untuk yg aktif saja
            DTA.FillBy_Year(TBL, ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi))

            For Each row In TBL.Rows
                If row.no_urut_max < row.ending_number Then
                    If row.no_urut_max = 0 Then
                        row.no_urut_max = row.starting_number - 1
                    End If

                    StartingNumb = row.no_urut_max
                    NoFakturTersedia += row.ending_number - StartingNumb
                    'For i As Integer = StartingNumb To row.ending_number
                    '    NoFakturTersedia += 1
                    'Next
                End If
            Next

            Return NoFakturTersedia
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            TBL.Dispose()
        End Try
    End Function
End Class