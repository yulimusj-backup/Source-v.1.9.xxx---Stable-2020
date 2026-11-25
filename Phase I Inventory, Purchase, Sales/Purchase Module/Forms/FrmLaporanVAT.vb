Option Explicit On
Option Strict Off

Imports Microsoft.VisualBasic.FileIO
Public Class FrmLaporanVAT
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""

    Dim DS As New DseLaporanVAT1
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim TipeSR As Sales_Module.ClsTransaksiSalesReturn.EnumTipeSR = Sales_Module.ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
    Dim TipePR As ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Nota_Retur
    Dim UseTax As Boolean = True
    Dim DefPrecBase As Integer = ObjCur.GetDecimal_DefaultCurrency
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjLicense As New Core_Module.ClsLicense

    Dim DTAin As New DseLaporanVAT1TableAdapters.DataTable1TableAdapter
    Dim TBLin As New DseLaporanVAT1.DataTable1DataTable
    Dim DTAOut As New DseLaporanVAT1TableAdapters.DataTable2TableAdapter
    Dim TBLOut As New DseLaporanVAT1.DataTable2DataTable
    Dim ObjInv As New Sales_Module.ClsTransaksiSalesInvoice
    Dim ObjAOP As New ClsTransaksiAOP

    Public Function FormatNPWP(ByVal str As String) As String
        Dim _NPWP As String
        Try
            If String.IsNullOrEmpty(str) = False Then
                _NPWP = str.Replace(".", "")
                _NPWP = _NPWP.Replace("-", "")
                Return _NPWP
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Enum EnumInOut As Integer
        VATIn
        VATOut
        Both
        VATIn_Purchase
        VATIn_SalesReturn
        VATOut_Sales
        VATOut_PurchaseReturn
    End Enum

    Public Enum EnumVATType
        VAT_IN_AOP
        VAT_IN_AOP_DP
        VAT_IN_QUICK_AOP
        VAT_IN_RETURN
        VAT_OUT_INVOICE
        VAT_OUT_DP
        VAT_OUT_QUICK_INVOICE
        VAT_OUT_RETURN
    End Enum

    Private Sub FrmLaporanVAT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.ComboBox1.SelectedIndex = 0
        Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreate.Click
        Dim ValTxt As String = "Kode Pajak;Kode Transaksi;Kode Status;Kode Dokumen;Flag VAT;NPWP / Nomor Paspor;Nama Lawan Transaksi;Nomor Faktur / Dokumen;Jenis Dokumen;Nomor Faktur Pengganti / Retur;Jenis Dokumen Dokumen Pengganti / Retur;Tanggal Faktur / Dokumen;Tanggal SSP;Masa Pajak;Tahun Pajak;Pembetulan;DPP;PPN;PPnBM"
        Dim RowLaporan As DseLaporanVAT1.Tbl_VATRow
        Dim Pajak As String = ""
        Dim retur As String = ""
        Dim DokPengganti As String = ""
        Dim KodeStatus As String = ""
        Dim jenisdok As Integer = 0
        Dim DummyIn As String = ""
        Dim DummyOut As String = ""
        Dim Dummy As String = ""

        Try
            'keluarkan datanya kembali
            Me.btnShow.PerformClick()

            'cek berkaitan dgn no NPWP 
            For Each RowLaporan In DS.Tbl_VAT.Rows
                If RowLaporan.VAT <> 0 Then
                    With RowLaporan
                        If .Type = "VAT IN" Then
                            If RowLaporan.NPWP = "" Then
                                If DummyIn = "" Then
                                    DummyIn = RowLaporan.Invoice_No
                                Else
                                    DummyIn = DummyIn + ", " + RowLaporan.Invoice_No
                                End If
                            End If
                        Else
                            If RowLaporan.NPWP = "" Then
                                If DummyOut = "" Then
                                    DummyOut = RowLaporan.Invoice_No
                                Else
                                    DummyOut = DummyOut + ", " + RowLaporan.Invoice_No
                                End If
                            End If
                        End If
                    End With
                End If
            Next
            If DummyIn <> "" Or DummyOut <> "" Then
                Dummy = "Please provide correct NPWP number for the following document:"
                If DummyIn <> "" Then
                    Dummy = Dummy + vbCrLf + "VAT IN: " + DummyIn
                End If
                If DummyOut <> "" Then
                    Dummy = Dummy + vbCrLf + "VAT OUT: " + DummyOut
                End If

                MsgBox(Dummy, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'cek format tax invoice number yang berkaitan pembelian dan penjualan
            Dummy = ""
            DummyIn = ""
            DummyOut = ""
            For Each RowLaporan In DS.Tbl_VAT.Rows
                With RowLaporan
                    If RowLaporan.VAT <> 0 Then
                        If RowLaporan.Type = "VAT IN" Then
                            If RowLaporan.Tax_Invoice_No = "" Or RowLaporan.Tax_Invoice_No.Length < 2 Then
                                If DummyIn = "" Then
                                    DummyIn = RowLaporan.Invoice_No
                                Else
                                    DummyIn = DummyIn + ", " + RowLaporan.Invoice_No
                                End If
                            End If
                        Else
                            If RowLaporan.Tax_Invoice_No = "" Or RowLaporan.Tax_Invoice_No.Length < 2 Then
                                If DummyOut = "" Then
                                    DummyOut = RowLaporan.Invoice_No
                                Else
                                    DummyOut = DummyOut + ", " + RowLaporan.Invoice_No
                                End If
                            End If
                        End If
                    End If
                End With
            Next
            If DummyIn <> "" Or DummyOut <> "" Then
                Dummy = "Please provide correct tax invoice number for the following document:"
                If DummyIn <> "" Then
                    Dummy = Dummy + vbCrLf + "VAT IN: " + DummyIn
                End If
                If DummyOut <> "" Then
                    Dummy = Dummy + vbCrLf + "VAT OUT: " + DummyOut
                End If

                MsgBox(Dummy, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If DS.Tbl_VAT.Rows.Count > 0 Then
                If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    For Each RowLaporan In DS.Tbl_VAT.Rows
                        With RowLaporan
                            If .Type = "VAT IN" Then
                                Pajak = "B"
                            Else
                                Pajak = "A"
                            End If
                            If .Retur = 0 Then
                                retur = ";1;"
                                DokPengganti = ""
                                jenisdok = 0
                            Else : .Retur = 1
                                retur = ";2;"
                                DokPengganti = .Tax_Invoice_No.ToString
                                jenisdok = 1
                            End If
                            If .Tax_Invoice_No = "" Or .Tax_Invoice_No.Length < 2 Then
                                KodeStatus = " "
                            Else
                                KodeStatus = .Tax_Invoice_No.ToString.Substring(1, 1)
                            End If
                            If .Type <> "VAT IN" And ._Date < #4/1/2013# And .Retur = 0 Then
                                .Tax_Invoice_No = .Tax_Invoice_No.Substring(0, 7) + "-" + .Tax_Invoice_No.Substring(7)
                            End If

                            If CDate(._Date).Month < 10 Then
                                ValTxt = ValTxt + vbCrLf + Pajak + ";2;" + KodeStatus + retur + "0;" + Me.FormatNPWP(.NPWP).ToString + ";" + .Cust_Supp_Name + ";" + .Tax_Invoice_No + ";" + jenisdok.ToString + ";" + DokPengganti + "; ;" + Format(CDate(._Date), "dd/MM/yyyy") + ";" + _
                                ";0".ToString + CDate(._Date).Month.ToString + "0".ToString + CDate(._Date).Month.ToString + ";" + CDate(._Date).Year.ToString + ";0;" + CStr(Math.Floor(CDbl(.Purchase_Sales))) + ";" + CStr(Math.Floor(CDbl(.VAT))) + ";0"
                            Else
                                ValTxt = ValTxt + vbCrLf + Pajak + ";2;" + KodeStatus + retur + "0;" + Me.FormatNPWP(.NPWP).ToString + ";" + .Cust_Supp_Name + ";" + .Tax_Invoice_No + ";" + jenisdok.ToString + ";" + DokPengganti + "; ;" + Format(CDate(._Date), "dd/MM/yyyy") + ";" + _
                                ";".ToString + CDate(._Date).Month.ToString + "".ToString + CDate(._Date).Month.ToString + ";" + CDate(._Date).Year.ToString + ";0;" + CStr(Math.Floor(CDbl(.Purchase_Sales))) + ";" + CStr(Math.Floor(CDbl(.VAT))) + ";0"
                            End If
                        End With
                    Next
                    My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, ValTxt, False)
                End If
            End If

            MsgBox("Csv file created.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim TypeInOut As Integer = 1
        Dim SalerOrPr As Decimal = 0
        Dim VAT As Decimal = 0

        Dim RowIn As DseLaporanVAT1.DataTable1Row
        Dim RowOut As DseLaporanVAT1.DataTable2Row

        Dim RowDgv As DataGridViewRow
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
           
            Me.dgvItem.Rows.Clear()
            If Me.ComboBox1.SelectedIndex = 0 Then
                TypeInOut = FrmLaporanVAT.EnumInOut.VATIn
            ElseIf Me.ComboBox1.SelectedIndex = 1 Then
                TypeInOut = FrmLaporanVAT.EnumInOut.VATOut
            Else
                TypeInOut = 3
            End If

            If TypeInOut = FrmLaporanVAT.EnumInOut.VATIn Then
                Me.DTAin.Fill(TBLin, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipePR)
                For Each RowIn In TBLin.Rows
                    With RowIn
                        If .VAT_dipungut = 0 Then
                            SalerOrPr = 0
                            VAT = 0
                        Else : .VAT_dipungut = 1
                            SalerOrPr = .Purchase
                            VAT = .VAT
                        End If
                        Me.dgvItem.Rows.Add()
                        RowDgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        RowDgv.Cells(Me.ColTipeTransaksiString.Index).Value = ._In
                        RowDgv.Cells(Me.ColSuppCust.Index).Value = .Supp
                        RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                        If .Retur = 0 Then
                            RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no2
                        Else
                            RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no
                        End If
                        RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .Tax_invoice_no
                        RowDgv.Cells(Me.ColDate.Index).Value = Format(.Tanggal_dari.Date, "dd MMM yyyy")
                        RowDgv.Cells(Me.ColPurchaseSales.Index).Value = Me.ObjCD.DigitGrouping(SalerOrPr.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(VAT.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColPrimaryDocAcuan.Index).Value = .Primary
                        RowDgv.Cells(Me.ColEnumType.Index).Value = .Retur
                        RowDgv.Cells(Me.ColEnumInOut.Index).Value = FrmLaporanVAT.EnumInOut.VATIn
                    End With
                Next
                Me.SalinCrv()

            ElseIf TypeInOut = EnumInOut.VATOut Then
                Me.DTAOut.Fill(TBLOut, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipeSR)
                For Each RowOut In Me.TBLOut
                    With RowOut
                        If .VAT_Dipungut = 0 Then
                            SalerOrPr = 0
                            VAT = 0
                        Else : .VAT_Dipungut = 1
                            SalerOrPr = .Sales
                            VAT = .VAT
                        End If
                        Me.dgvItem.Rows.Add()
                        RowDgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        RowDgv.Cells(Me.ColTipeTransaksiString.Index).Value = .Out
                        RowDgv.Cells(Me.ColSuppCust.Index).Value = .Cust
                        RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                        RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no
                        RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .No_Faktur_Pajak
                        RowDgv.Cells(Me.ColDate.Index).Value = Format(.Tanggal_dari.Date, "dd MMM yyyy")
                        RowDgv.Cells(Me.ColPurchaseSales.Index).Value = Me.ObjCD.DigitGrouping(SalerOrPr.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(VAT.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColPrimaryDocAcuan.Index).Value = .Primary
                        RowDgv.Cells(Me.ColEnumType.Index).Value = .Retur
                        RowDgv.Cells(Me.ColEnumInOut.Index).Value = FrmLaporanVAT.EnumInOut.VATOut
                    End With
                Next
                Me.SalinCrv()
            ElseIf TypeInOut = 3 Then
                Me.DTAin.Fill(TBLin, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipePR)
                For Each RowIn In TBLin.Rows
                    With RowIn
                        If .VAT_dipungut = 0 Then
                            SalerOrPr = 0
                            VAT = 0
                        Else : .VAT_dipungut = 1
                            SalerOrPr = .Purchase
                            VAT = .VAT
                        End If
                        Me.dgvItem.Rows.Add()
                        RowDgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        RowDgv.Cells(Me.ColTipeTransaksiString.Index).Value = ._In
                        RowDgv.Cells(Me.ColSuppCust.Index).Value = .Supp
                        RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                        If .Retur = 0 Then
                            RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no2
                        Else
                            RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no
                        End If
                        RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .Tax_invoice_no
                        RowDgv.Cells(Me.ColDate.Index).Value = Format(.Tanggal_dari.Date, "dd MMM yyyy")
                        RowDgv.Cells(Me.ColPurchaseSales.Index).Value = Me.ObjCD.DigitGrouping(SalerOrPr.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(VAT.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColPrimaryDocAcuan.Index).Value = .Primary
                        RowDgv.Cells(Me.ColEnumType.Index).Value = .Retur
                        RowDgv.Cells(Me.ColEnumInOut.Index).Value = FrmLaporanVAT.EnumInOut.VATIn
                    End With
                Next
                Me.DTAOut.Fill(TBLOut, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipeSR)
                For Each RowOut In Me.TBLOut
                    With RowOut
                        If .VAT_Dipungut = 0 Then
                            SalerOrPr = 0
                            VAT = 0
                        Else : .VAT_Dipungut = 1
                            SalerOrPr = .Sales
                            VAT = .VAT
                        End If
                        Me.dgvItem.Rows.Add()
                        RowDgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        RowDgv.Cells(Me.ColTipeTransaksiString.Index).Value = .Out
                        RowDgv.Cells(Me.ColSuppCust.Index).Value = .Cust
                        RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                        RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no
                        RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .No_Faktur_Pajak
                        RowDgv.Cells(Me.ColDate.Index).Value = Format(.Tanggal_dari.Date, "dd MMM yyyy")
                        RowDgv.Cells(Me.ColPurchaseSales.Index).Value = Me.ObjCD.DigitGrouping(SalerOrPr.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(VAT.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColPrimaryDocAcuan.Index).Value = .Primary
                        RowDgv.Cells(Me.ColEnumType.Index).Value = .Retur
                        RowDgv.Cells(Me.ColEnumInOut.Index).Value = FrmLaporanVAT.EnumInOut.VATOut
                    End With
                Next
                Me.SalinCrv()
            End If

            Me.crv1.ShowExportButton = True
            Me.crv1.ShowGotoPageButton = True
            Me.crv1.ShowPageNavigateButtons = True
            Me.crv1.ShowPrintButton = True
            Me.crv1.ShowTextSearchButton = True
            Me.crv1.ShowZoomButton = True
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub SalinCrv()
        Dim rpt As New CrpLaporanVAT
        Dim TBLLap As DseLaporanVAT1.Tbl_VATDataTable = DS.Tbl_VAT
        Dim RowLap As DseLaporanVAT1.Tbl_VATRow
        Dim RowDGV As DataGridViewRow
        Dim InOut As Integer = 0

        Try
            Me.DS.Tbl_VAT.Rows.Clear()
            For Each RowDGV In Me.dgvItem.Rows
                With RowDGV
                    RowLap = TBLLap.NewRow
                    RowLap.Type = .Cells(Me.ColTipeTransaksiString.Index).Value
                    RowLap.Cust_Supp_Name = .Cells(Me.ColSuppCust.Index).Value
                    RowLap.NPWP = .Cells(Me.ColNPWPNumber.Index).Value
                    RowLap.Invoice_No = .Cells(Me.ColInvoiceNumber.Index).Value
                    RowLap.Tax_Invoice_No = .Cells(Me.ColTaxInvoiceNumber.Index).Value
                    RowLap._Date = .Cells(Me.ColDate.Index).Value
                    RowLap.Purchase_Sales = .Cells(Me.ColPurchaseSales.Index).Value
                    RowLap.VAT = .Cells(Me.ColVAT.Index).Value
                    RowLap.Retur = .Cells(Me.ColEnumType.Index).Value
                    TBLLap.AddTbl_VATRow(RowLap)
                End With
            Next

            rpt.SetDataSource(DS)
            rpt.SetParameterValue(rpt.Parameter_Tanggal_Dari.ParameterFieldName, Format(Me.dtpFrom.Value.Date, "dd MMM yyyy"))
            rpt.SetParameterValue(rpt.Parameter_Tanggal_Sampai.ParameterFieldName, Format(Me.dtpTo.Value.Date, "dd MMM yyyy"))

            crv1.ReportSource = rpt
            crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub dgvItem_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellClick
        Dim RowDGV As DataGridViewRow
        Dim InOut As FrmLaporanVAT.EnumInOut = EnumInOut.VATIn
        Dim PrimaryDoc As Double = -1

        Try
            If e.ColumnIndex = Me.ColInvoiceNumber.Index Then
                RowDGV = Me.dgvItem.Rows(e.RowIndex)
                InOut = RowDGV.Cells(Me.ColEnumInOut.Index).Value
                PrimaryDoc = RowDGV.Cells(Me.ColPrimaryDocAcuan.Index).Value

                If InOut = EnumInOut.VATIn Then
                    If RowDGV.Cells(Me.ColEnumType.Index).Value = 0 Then
                        'aop
                        If Me.ObjAOP.ApakahQuickPurchase(PrimaryDoc) = True Then
                            Me.Show_QuickAOP(PrimaryDoc)
                        ElseIf Me.ObjAOP.ApakahAOPDP(PrimaryDoc) = True Then
                            Me.Show_AOPDP(PrimaryDoc)
                        Else
                            Me.Show_AOP(PrimaryDoc)
                        End If
                    Else
                        'purchase return 
                        Me.Show_PR(PrimaryDoc)
                    End If
                Else
                    If RowDGV.Cells(Me.ColEnumType.Index).Value = 0 Then
                        'invoice
                        If Me.ObjInv.ApakahQuickInvoice(PrimaryDoc) = True Then
                            Me.Show_QuickInvoice(PrimaryDoc)
                        ElseIf Me.ObjInv.ApakahInvoiceDP(PrimaryDoc) = True Then
                            Me.Show_InvoiceDP(PrimaryDoc)
                        Else
                            Me.Show_Invoice(PrimaryDoc)
                        End If
                    Else
                        'sr
                        Me.Show_SR(PrimaryDoc)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Show_PR(ByVal vPrimary As Double)
        Try
            'cek apakah user memiliki hak akses yang cukup
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiPurchaseReturn") = False Then
                MsgBox("Insufficient acces", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Dim ij As Purchase_Module.FrmTransaksiPurchaseReturn
            Dim j As Windows.Forms.Form

            For Each j In Me.ParentForm.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPurchaseReturn)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    ij.ShowTransaction(vPrimary)
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiPurchaseReturn
            ' Increment the caption counter.
            i.UserCode = Me.UserCode.ToString
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me.ParentForm
            i.Show()
            i.ShowTransaction(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Show_AOP(ByVal vPrimary As Double)
        Try
            'cek apakah user memiliki hak akses yang cukup
            If Me.ObjAOP.ApakahLokal(vPrimary) = True Then
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiAOPLokal") = False Then
                    MsgBox("Insufficient acces", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                Dim ij As Purchase_Module.FrmTransaksiAOP2
                Dim j As Windows.Forms.Form

                For Each j In Me.ParentForm.MdiChildren
                    ij = TryCast(j, Purchase_Module.FrmTransaksiAOP2)
                    If IsNothing(ij) = False Then
                        If ij.Lokal = True Then
                            ij.BringToFront()
                            ij.Focus()
                            ij.ShowTransaction(vPrimary)
                            Exit Sub
                        End If
                    End If
                Next

                Dim i As New Purchase_Module.FrmTransaksiAOP2
                ' Increment the caption counter.
                i.UserCode = Me.UserCode.ToString
                ' Set the caption to be unique.
                i.BackColorAll = Core_Module.GetSMESSBackColorAll
                i.MdiParent = Me.ParentForm
                i.Lokal = True
                i.Show()
                i.ShowTransaction(vPrimary)
            Else
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiAOPRImport") = False Then
                    MsgBox("Insufficient acces", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                Dim ij As Purchase_Module.FrmTransaksiAOP2
                Dim j As Windows.Forms.Form

                For Each j In Me.ParentForm.MdiChildren
                    ij = TryCast(j, Purchase_Module.FrmTransaksiAOP2)
                    If IsNothing(ij) = False Then
                        If ij.Lokal = False Then
                            ij.BringToFront()
                            ij.Focus()
                            ij.ShowTransaction(vPrimary)
                            Exit Sub
                        End If
                    End If
                Next

                Dim i As New Purchase_Module.FrmTransaksiAOP2
                ' Increment the caption counter.
                i.UserCode = Me.UserCode.ToString
                ' Set the caption to be unique.
                i.BackColorAll = Core_Module.GetSMESSBackColorAll
                i.MdiParent = Me.ParentForm
                i.Lokal = False
                i.Show()
                i.ShowTransaction(vPrimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Show_AOPDP(ByVal vPrimary As Double)
        Try
            'cek apakah user memiliki hak akses yang cukup
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiAOPDP") = False Then
                MsgBox("Insufficient acces", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Dim ij As Purchase_Module.FrmTransaksiAOPDP
            Dim j As Windows.Forms.Form

            For Each j In Me.ParentForm.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiAOPDP)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    ij.ShowTransaction(vPrimary)
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiAOPDP
            ' Increment the caption counter.
            i.UserCode = Me.UserCode.ToString
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me.ParentForm
            i.Show()
            i.ShowTransaction(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Show_QuickAOP(ByVal vPrimary As Double)
        Try
            'cek apakah user memiliki hak akses yang cukup
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiAOP_Quick") = False Then
                MsgBox("Insufficient acces", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Dim ij As Purchase_Module.FrmTransaksiAOP_Quick2
            Dim j As Windows.Forms.Form

            For Each j In Me.ParentForm.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiAOP_Quick2)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    ij.ShowTransaction(vPrimary)
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiAOP_Quick2
            ' Increment the caption counter.
            i.UserCode = Me.UserCode.ToString
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me.ParentForm
            i.Show()
            i.ShowTransaction(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Show_SR(ByVal vPrimary As Double)
        Try
            'cek apakah user memiliki hak akses yang cukup
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSalesReturn") = False Then
                MsgBox("Insufficient acces", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Dim ij As Sales_Module.FrmTransaksiSalesReturn
            Dim j As Windows.Forms.Form

            For Each j In Me.ParentForm.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiSalesReturn)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    ij.ShowTransaction(vPrimary)
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiSalesReturn
            ' Increment the caption counter.
            i.UserCode = Me.UserCode.ToString
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me.ParentForm
            i.Show()
            i.ShowTransaction(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Show_QuickInvoice(ByVal vPrimary As Double)
        Try
            'cek apakah user memiliki hak akses yang cukup
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSalesInvoiceQuick") = False Then
                MsgBox("Insufficient acces", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Dim ij As Sales_Module.FrmTransaksiSalesInvoiceQuick
            Dim j As Windows.Forms.Form

            For Each j In Me.ParentForm.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiSalesInvoiceQuick)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    ij.ShowTransaction(vPrimary)
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiSalesInvoiceQuick
            ' Increment the caption counter.
            i.UserCode = Me.UserCode.ToString
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me.ParentForm
            i.Show()
            i.ShowTransaction(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Show_Invoice(ByVal vPrimary As Double)
        Try
            'cek apakah user memiliki hak akses yang cukup
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSalesInvoice") = False Then
                MsgBox("Insufficient acces", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Dim ij As Sales_Module.FrmTransaksiSalesInvoice
            Dim j As Windows.Forms.Form

            For Each j In Me.ParentForm.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiSalesInvoice)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    ij.ShowTransaction(vPrimary)
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiSalesInvoice
            ' Increment the caption counter.
            i.UserCode = Me.UserCode.ToString
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me.ParentForm
            i.Show()
            i.ShowTransaction(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Show_InvoiceDP(ByVal vPrimary As Double)
        Try
            'cek apakah user memiliki hak akses yang cukup
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSalesInvoiceDP") = False Then
                MsgBox("Insufficient acces", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Dim ij As Sales_Module.FrmTransaksiSalesInvoiceDP
            Dim j As Windows.Forms.Form

            For Each j In Me.ParentForm.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiSalesInvoiceDP)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    ij.ShowTransaction(vPrimary)
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiSalesInvoiceDP
            ' Increment the caption counter.
            i.UserCode = Me.UserCode.ToString
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me.ParentForm
            i.Show()
            i.ShowTransaction(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCreateDataImpor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateDataImpor.Click
        Dim ValTxt As String = "Primary;No Transaksi;Supplier;NPWP;Tax Invoice"
        Dim DTA As New DseLaporanVAT1TableAdapters.DataTable3TableAdapter
        Dim TBL As New DseLaporanVAT1.DataTable3DataTable
        Dim row As DseLaporanVAT1.DataTable3Row
        Dim DTA_Supplier As New DseLaporanVAT1TableAdapters.tblmaster_supplierTableAdapter
        Dim TBL_Supplier As New DseLaporanVAT1.tblmaster_supplierDataTable
        Dim row_Supplier As DseLaporanVAT1.tblmaster_supplierRow

        Try
            'isikan data
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)

            'untuk no npwp yang kosong, cek ke master supplier
            DTA_Supplier.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Supplier.Fill(TBL_Supplier)
            For Each row In TBL.Rows
                With row
                    If .No_NPWP = "" Then
                        TBL_Supplier.DefaultView.RowFilter = "[Supplier Name] = '" + .Supplier_Name + "'"
                        row_Supplier = TBL_Supplier.DefaultView(0).Row
                        .No_NPWP = row_Supplier.NPWP_No
                    End If
                End With
            Next

            'export
            For Each row In TBL.Rows
                With row
                    ValTxt = ValTxt + vbCrLf + .Primary.ToString + ";" + .Supplier_Invoice_No + ";" + .Supplier_Name + ";" + .No_NPWP + ";" + .Tax_Invoice_Number
                End With
            Next
            If TBL.Rows.Count > 0 Then
                Me.SaveFileDialog1.Title = "Export AOP Data"
                If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, ValTxt, False)
                End If
            End If

            MsgBox("Data export completed", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA_Supplier.Dispose()
            TBL.Dispose()
            TBL_Supplier.Dispose()
        End Try
    End Sub

    Private Sub btnImportData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImportData.Click
        Dim i As Integer = 0
        Dim PrimaryAOP As Double = -1
        Dim NPWP_Supplier As String = ""
        Dim TaxInvoiceNumber As String = ""
        Dim DTA_AOP As New DseLaporanVAT1TableAdapters.DataTable3TableAdapter

        Try
            DTA_AOP.Connection.ConnectionString = Core_Module.GetConstring

            If Me.OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                If Me.OpenFileDialog1.FileName <> "" Then
                    Dim filename As String = Me.OpenFileDialog1.FileName
                    Dim fields As String()
                    Dim delimiter As String = ";"
                    Using parser As New TextFieldParser(filename)
                        parser.SetDelimiters(delimiter)
                        While Not parser.EndOfData
                            ' Read in the fields for the current line
                            fields = parser.ReadFields()
                            ' Add code here to use data in fields variable.
                            If i > 0 Then
                                'update ke dalam database aop
                                NPWP_Supplier = fields(3)
                                PrimaryAOP = fields(0)
                                TaxInvoiceNumber = fields(1)

                                DTA_AOP.UpdateQuery(TaxInvoiceNumber, NPWP_Supplier, PrimaryAOP)
                            End If
                            'ke baris selanjutnya
                            i += 1
                        End While
                    End Using

                    MsgBox("AOP data updated", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_AOP.Dispose()
        End Try
    End Sub
End Class