''inf-001
Option Explicit On
Option Strict Off

Imports System.Data
Imports Microsoft.VisualBasic.FileIO

Public Class FrmLaporanVAT2025
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""

    Dim DS As New DseLaporanVAT2025
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency
    'Dim TipeSR As Sales_Module.ClsTransaksiSalesReturn.EnumTipeSR = Sales_Module.ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
    Dim TipePR As ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Nota_Retur
    Dim UseTax As Boolean = True
    Dim DefPrecBase As Integer = ObjCur.GetDecimal_DefaultCurrency
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjPR As New ClsTransaksiPurchaseReturn

    Dim DTAin As New DseLaporanVAT2025TableAdapters.DataTable1TableAdapter
    Dim TBLin As New DseLaporanVAT2025.DataTable1DataTable
    Dim DTAIN_Purchase As New DseLaporanVAT2025TableAdapters.TBL_AOPTableAdapter
    Dim TBLIN_Purchase As New DseLaporanVAT2025.TBL_AOPDataTable
    Dim rowIN_Purchase As DseLaporanVAT2025.TBL_AOPRow
    Dim DTAIN_SR As New DseLaporanVAT2025TableAdapters.Sales_ReturnTableAdapter
    Dim TBLIN_SR As New DseLaporanVAT2025.Sales_ReturnDataTable
    Dim rowIN_SR As DseLaporanVAT2025.Sales_ReturnRow

    Dim DTAOut As New DseLaporanVAT2025TableAdapters.DataTable2TableAdapter
    Dim TBLOut As New DseLaporanVAT2025.DataTable2DataTable
    Dim DTAOut_Sales As New DseLaporanVAT2025TableAdapters.InvoiceTableAdapter
    Dim TBLOut_Sales As New DseLaporanVAT2025.InvoiceDataTable
    Dim rowOut_Sales As DseLaporanVAT2025.InvoiceRow
    Dim DTAReturn As New DseLaporanVAT2025TableAdapters.Item_Invoice_ReturnTableAdapter
    Dim TBLReturn As New DseLaporanVAT2025.Item_Invoice_ReturnDataTable
    Dim rowReturn As DseLaporanVAT2025.Item_Invoice_ReturnRow
    Dim DTA As New DseLaporanVAT2025TableAdapters.Item_InvoiceTableAdapter
    Dim TBL As New DseLaporanVAT2025.Item_InvoiceDataTable
    Dim row As DseLaporanVAT2025.Item_InvoiceRow
    Dim DTAItemDP As New DseLaporanVAT2025TableAdapters.Item_Invoice_DPTableAdapter
    Dim TBLItemDP As New DseLaporanVAT2025.Item_Invoice_DPDataTable
    Dim rowItemDP As DseLaporanVAT2025.Item_Invoice_DPRow

    Dim DTAOut_PR As New DseLaporanVAT2025TableAdapters.Purchase_ReturnTableAdapter
    Dim TBLOut_PR As New DseLaporanVAT2025.Purchase_ReturnDataTable
    Dim rowOut_PR As DseLaporanVAT2025.Purchase_ReturnRow
    Dim DTAOut_PR_Ref As New DseLaporanVAT2025TableAdapters.Purchase_Return_ReferenceTableAdapter
    Dim TBLOut_PR_Ref As New DseLaporanVAT2025.Purchase_Return_ReferenceDataTable
    Dim rowOut_PR_Ref As DseLaporanVAT2025.Purchase_Return_ReferenceRow
    Dim ObjInv As New Sales_Module.ClsTransaksiSalesInvoice
    'Dim ObjInv As New ClsTransaksiSalesInvoice
    Dim ObjAOP As New ClsTransaksiAOP
    Dim ObjSR As New Sales_Module.ClsTransaksiSalesReturn
    Dim OptionNIKPASS As Core_Module.clsMasterCustomer
    Dim ObjCustomer As New Core_Module.clsMasterCustomer
    Dim ObjLicense As New Core_Module.ClsLicense
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting

    '~ab~
    'Remarks : tambah hak akses bisa edit tax number
    '03/02/2020
    Dim BisaEditTaxNo As Boolean = False
    Dim BufferTax As String = ""
    '~

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
    End Enum

    Public Enum EnumVATType
        VAT_OUT_INVOICE
        VAT_OUT_INVOICEDP
    End Enum

    Private Sub FrmLaporanVAT2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.ChkShowDiscount.Checked = True Then
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.VAT, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ShowDiskon, "1")
            Else
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.VAT, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ShowDiskon, "0")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanVAT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Dummy As String = ""

        Try
            Me.ComboBox1.SelectedIndex = 0
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTAin.Connection.ConnectionString = DTA.Connection.ConnectionString
            Me.DTAIN_Purchase.Connection.ConnectionString = DTA.Connection.ConnectionString
            Me.DTAIN_SR.Connection.ConnectionString = DTA.Connection.ConnectionString
            Me.DTAOut.Connection.ConnectionString = DTA.Connection.ConnectionString
            Me.DTAOut_PR.Connection.ConnectionString = DTA.Connection.ConnectionString
            Me.DTAOut_PR_Ref.Connection.ConnectionString = DTA.Connection.ConnectionString
            Me.DTAOut_Sales.Connection.ConnectionString = DTA.Connection.ConnectionString
            Me.DTAReturn.Connection.ConnectionString = DTA.Connection.ConnectionString
            Me.DTAItemDP.Connection.ConnectionString = DTA.Connection.ConnectionString
            Me.ColTipeTransaksiString.Visible = False

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.VAT, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ShowDiskon)
            If Dummy = "1" Then
                Me.ChkShowDiscount.Checked = True
            End If
            Me.dtpFrom.Value = Me.ObjCD.GetFirstDateofTheMonth(Now)
            Me.dtpTo.Value = Me.ObjCD.GetLasDateOfTheMonth(Now)

            '~ab~
            'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
            '03/02/2020
            Me.dgvItem.ReadOnly = False
            '~
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreate.Click
        Dim ValTxt As String = "Kode Pajak;Kode Transaksi;Kode Status;Kode Dokumen;Flag VAT;NPWP / Nomor Paspor;Nama Lawan Transaksi;Nomor Faktur / Dokumen;Jenis Dokumen;Nomor Faktur Pengganti / Retur;Jenis Dokumen Dokumen Pengganti / Retur;Tanggal Faktur / Dokumen;Tanggal SSP;Masa Pajak;Tahun Pajak;Pembetulan;DPP;PPN;PPnBM"
        Dim RowLaporan As DseLaporanVAT2025.Tbl_VATRow
        Dim Pajak As String = ""
        Dim retur As String = ""
        Dim DokPengganti As String = ""
        Dim KodeStatus As String = ""
        Dim jenisdok As Integer = 0
        Dim DummyIn As String = ""
        Dim DummyOut As String = ""
        Dim Dummy As String = ""

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
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

            MsgBox("Csv file created. Please be aware that characters used in import process (;,"") will be omitted.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ExportPajakMasukan()
        Dim ValTxt As String = "FM;KD_JENIS_TRANSAKSI;FG_PENGGANTI;NOMOR_FAKTUR;MASA_PAJAK;TAHUN_PAJAK;TANGGAL_FAKTUR;NPWP;NAMA;ALAMAT_LENGKAP;JUMLAH_DPP;JUMLAH_PPN;JUMLAH_PPNBM;IS_CREDITABLE"
        Dim RowLaporan As DseLaporanVAT2025.Tbl_VATRow
        Dim DummyIn As String = ""
        Dim DummyOut As String = ""
        Dim Dummy As String = ""

        Dim NoFakturDummy As String = ""
        Dim NoNPWP As String = ""

        Try
            'keluarkan datanya kembali
            'Me.btnShow.PerformClick()

            'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
            '03/02/2020 
            Dim tblDummy As New DseLaporanVAT2025.Tbl_VATDataTable
            Dim rowDummy As DseLaporanVAT2025.Tbl_VATRow

            Dim ListExport As New List(Of Double)
            Dim ListExportStatus As New List(Of Double)
            ListExport.Clear()
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If row.Cells(Me.ColExport.Index).Value = Boolean.TrueString Then 'jika betul2 belum pernah di export
                    If ListExport.Contains(row.Cells(Me.ColPrimaryDocAcuan.Index).Value) = False Then
                        ListExport.Add(row.Cells(Me.ColPrimaryDocAcuan.Index).Value)
                    End If
                Else
                    If String.IsNullOrEmpty(row.Cells(Me.ColExportStatus.Index).Value) = False Then ' jika sudah pernah di export
                        If ListExportStatus.Contains(row.Cells(Me.ColPrimaryDocAcuan.Index).Value) = False Then
                            ListExportStatus.Add(row.Cells(Me.ColPrimaryDocAcuan.Index).Value)
                        End If
                    End If
                End If
            Next
            '~

            'cek berkaitan dgn no NPWP 
            For Each RowLaporan In DS.Tbl_VAT.Rows
                If ListExport.Contains(RowLaporan.primary_sales_invoice) = True Then
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
                End If
            Next

            If DummyIn <> "" Or DummyOut <> "" Then
                Dummy = "Please provide correct NPWP number for the following document:"
                If DummyIn <> "" Then
                    Dummy = Dummy + vbCrLf + DummyIn
                End If
                If DummyOut <> "" Then
                    Dummy = Dummy + vbCrLf + DummyOut
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
                    If ListExport.Contains(RowLaporan.primary_sales_invoice) = True Then
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
                    End If
                End With
            Next

            If DummyIn <> "" Or DummyOut <> "" Then
                Dummy = "Please provide correct tax invoice number for the following document:"
                If DummyIn <> "" Then
                    Dummy = Dummy + vbCrLf + DummyIn
                End If
                If DummyOut <> "" Then
                    Dummy = Dummy + vbCrLf + DummyOut
                End If

                MsgBox(Dummy, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '~ab~
            'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
            '03/02/2020
            If ListExport.Count = 0 Then
                MsgBox("Please check export minimum one item.", MsgBoxStyle.Information, "Information")
                Me.btnShow.PerformClick()
                Exit Sub
            End If

            tblDummy.Rows.Clear()
            tblDummy = DS.Tbl_VAT.Copy
            For i As Integer = tblDummy.Rows.Count - 1 To 0 Step -1
                rowDummy = tblDummy.Rows(i)
                If ListExport.Contains(rowDummy.primary_sales_invoice) = False Then
                    tblDummy.Rows.Remove(rowDummy)
                End If
            Next

            'If DS.Tbl_VAT.Rows.Count > 0 Then
            If tblDummy.Rows.Count > 0 Then
                If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor

                    'For Each RowLaporan In DS.Tbl_VAT.Rows
                    For Each rowDummy In tblDummy.Rows
                        With rowDummy
                            NoFakturDummy = .Tax_Invoice_No
                            NoFakturDummy = NoFakturDummy.Trim
                            NoFakturDummy = NoFakturDummy.Replace("-", "")
                            NoFakturDummy = NoFakturDummy.Replace(".", "")
                            NoNPWP = .NPWP
                            NoNPWP = NoNPWP.Trim
                            NoNPWP = NoNPWP.Replace("-", "")
                            NoNPWP = NoNPWP.Replace(".", "")

                            .Cust_Supp_Name = .Cust_Supp_Name.Replace(";", "")
                            .Invoice_No = .Invoice_No.Replace(";", "")
                            .NPWP = .NPWP.Replace(";", "")
                            .Retur = .Retur.Replace(";", "")
                            .Tax_Address = .Tax_Address.Replace(";", "")
                            .Tax_Invoice_No = .Tax_Invoice_No.Replace(";", "")
                            .Type = .Type.Replace(";", "")

                            .Cust_Supp_Name = .Cust_Supp_Name.Replace("""", "")
                            .Invoice_No = .Invoice_No.Replace("""", "")
                            .NPWP = .NPWP.Replace("""", "")
                            .Retur = .Retur.Replace("""", "")
                            .Tax_Address = .Tax_Address.Replace("""", "")
                            .Tax_Invoice_No = .Tax_Invoice_No.Replace("""", "")
                            .Type = .Type.Replace("""", "")

                            ValTxt = ValTxt + vbCrLf + "FM;" + NoFakturDummy.Substring(0, 2).ToCharArray + ";" + NoFakturDummy.Substring(2, 1) + ";" + NoFakturDummy.Substring(3) + ";" _
                            + CDate(._Date).Month.ToString + ";" + CDate(._Date).Year.ToString + ";" + Format(CDate(._Date), "dd/MM/yyyy") + ";" _
                            + NoNPWP + ";" + .Cust_Supp_Name + ";" + .Tax_Address + ";" + CStr(Math.Floor(CDbl(.Purchase_Sales))) + ";" + CStr(Math.Floor(CDbl(.VAT))) + ";0;1"

                        End With
                    Next

                    My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, ValTxt, False)
                    MsgBox("Csv file created. Please be aware that characters used in import process (;,"") will be omitted.", MsgBoxStyle.Information, "Information")

                    For Each RowLaporan In DS.Tbl_VAT.Rows
                        '~ab~
                        'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
                        '03/02/2020
                        With RowLaporan
                            If ListExport.Contains(.primary_sales_invoice) = True Then
                                Me.UpdateExportStatus_InPurchase(Me.UserCode & " on " & Format(Now, "dd MMM yyyy - HH.mm"), .primary_sales_invoice)
                            Else
                                If ListExportStatus.Contains(.primary_sales_invoice) = False Then
                                    Me.UpdateExportStatus_InPurchase("-", .primary_sales_invoice)
                                End If
                            End If
                        End With
                    Next
                End If
            End If
            '~ab~
            'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
            '03/02/2020
            Me.btnShow.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function MasukanRefLainnya(ByRef ValTxt As String, ByVal Tbl5 As DseLaporanVAT2025.DataTable5DataTable, ByVal Tbl6 As DseLaporanVAT2025.DataTable6DataTable, _
    ByVal InvNo As String, ByVal InvoiceDate As Date, ByVal CustomerName As String) As Boolean
        Dim Row5 As DseLaporanVAT2025.DataTable5Row
        Dim DONo As String = ""
        Dim SONo As String = "" '2024-12-17 IM, tambahan SO
        Dim Kurs As String = ""

        Dim Row6 As DseLaporanVAT2025.DataTable6Row
        Dim NoKMK As String = ""
        Dim TglKMK As String = ""
        Dim Txt As String = ""
        Dim DPP As String = ""
        Dim ObjIdentifier As New CustomSMESS.ClsIdentifier
        Dim DefPrecIDR As Decimal = Me.ObjCur.GetDecimal_DefaultCurrency()
        Dim PrimaryCust As Double = -1
        Try
            '~ab~
            'Remarks : penambahan hasil keluaran export ditambah no do, dpp, kurs & kmk no.
            '04/02/2020
            Tbl5.DefaultView.RowFilter = "[NoInvoice]='" & InvNo & "'"
            If Tbl5.DefaultView.Count > 0 Then
                For i As Integer = 0 To Tbl5.DefaultView.Count - 1
                    Row5 = Tbl5.DefaultView(i).Row
                    If DONo = "" Then
                        DONo = Row5.DO_No
                    Else
                        If DONo.Contains(Row5.DO_No) = False Then
                            DONo &= ", " & Row5.DO_No
                        End If
                    End If
                    '2024-12-17 IM, tambahan SO
                    If SONo = "" Then
                        SONo = Row5.SO_NO
                    Else
                        If SONo.Contains(Row5.SO_NO) = False Then
                            SONo &= ", " & Row5.SO_NO
                        End If
                    End If
                    If DPP = "" Then
                        DPP = Me.ObjCD.DigitGrouping(Row5.Sub_Total.ToString, 2) 'permintaan bu kusmi decimal di paksa 2 saja 'Row5.Default_Precision)
                    End If
                    If Kurs = "" And Row5.Primary_Currency <> 1 Then
                        Kurs = Me.ObjCD.DigitGrouping(Row5.Tax_Exchange_Rate.ToString, 0) 'permintaan bu kusmi decimal di paksa 0 saja 'DefPrecIDR)
                    End If
                    If PrimaryCust = -1 Then
                        PrimaryCust = Row5.Primary_Customer
                    End If
                Next
            End If

            Tbl6.DefaultView.RowFilter = "[start date] <= '" & Format(InvoiceDate, "dd MMM yyy") & "' and [end date] >= '" & Format(InvoiceDate, "dd MMM yyyy") & "'"
            If Tbl6.DefaultView.Count > 0 Then
                For i As Integer = 0 To Tbl6.DefaultView.Count - 1
                    Row6 = Tbl6.DefaultView(i).Row
                    If NoKMK = "" Then
                        NoKMK = Row6.kmk_number.ToString
                        TglKMK = Me.ObjCD.GetDateString_Indo_Panjang(Row6.kmk_date) 'permintaan bu kusmi untuk bulan nya ingin di tampilkan full 'Format(Row6.kmk_date, "dd MMM yyyy")
                        Exit For
                    End If
                Next
            End If

            If ObjIdentifier.GetIdentifier() = "smess_iii" Then
                If PrimaryCust = 416 Or PrimaryCust = 922 Then 'customer = PT MURAMOTO ELEKTRONIKA INDONESIA, dan PT. INDONESIA EPSON INDUSTRI
                    If DONo <> "" Then
                        Txt &= vbNewLine & "DO No : " & DONo
                        'Txt &= vbNewLine & "DPP : " & DPP
                    End If
                    '2024-12-17 IM, tambahan SO dan DPP dipisah panggilnya
                    If SONo <> "" Then
                        Txt &= vbNewLine & "SO No : " & SONo
                    End If
                    If DPP <> "" Then
                        Txt &= vbNewLine & "DPP : " & DPP
                    End If
                    'end 2024-12-17

                    If Kurs <> "" Then
                        Txt &= vbNewLine & "Kurs : Rp." & Kurs & ",-/USD 1"
                    End If
                    If NoKMK <> "" Then
                        Txt &= vbNewLine & "Berdasarkan KMK No. " & NoKMK & " " & TglKMK
                    End If
                End If

            ElseIf ObjIdentifier.GetIdentifier() = "smess_cls" Or ObjIdentifier.GetIdentifier() = "smess_asp2" Then
                If DONo <> "" Then
                    Txt &= vbNewLine & "DO No : " & DONo
                    Txt &= vbNewLine & "DPP : " & DPP
                End If
                If Kurs <> "" Then
                    Txt &= vbNewLine & "Kurs : Rp." & Kurs & ",-/USD 1"
                End If
                If NoKMK <> "" Then
                    Txt &= vbNewLine & "Berdasarkan KMK No. " & NoKMK & " " & TglKMK
                End If
            End If



            Txt = Txt & Chr(34)
            ValTxt = ValTxt & Txt
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub ExportPajakKeluaran()
        Dim ValTxt As String = "FK;KD_JENIS_TRANSAKSI;FG_PENGGANTI;NOMOR_FAKTUR;MASA_PAJAK;TAHUN_PAJAK;TANGGAL_FAKTUR;NPWP;NAMA;ALAMAT_LENGKAP;JUMLAH_DPP;JUMLAH_PPN;JUMLAH_PPNBM;ID_KETERANGAN_TAMBAHAN;FG_UANG_MUKA;UANG_MUKA_DPP;UANG_MUKA_PPN;UANG_MUKA_PPNBM;REFERENSI (NIK/PASSPORT);SPPB_NO"  '//TR : Tambahan baru Nik/Pas
        ValTxt = ValTxt + vbCrLf + "LT;NPWP;NAMA;JALAN;BLOK;NOMOR;RT;RW;KECAMATAN;KELURAHAN;KABUPATEN;PROPINSI;KODE_POS;NO_TELEPON"
        ValTxt = ValTxt + vbCrLf + "OF;KODE_OBJEK;NAMA;HARGA_SATUAN;JUMLAH_BARANG;HARGA_TOTAL;DISKON;DPP;PPN;TARIF_PPNBM;PPNBM"

        Dim RowLaporan As DseLaporanVAT2025.Tbl_VATRow
        Dim DummyIn As String = ""
        Dim DummyOut As String = ""
        Dim Dummy As String = ""
        Dim KodeJenisTransaksiPajakDummy As String = ""
        Dim NoFakturDummy As String = ""
        Dim NoNPWP As String = ""
        Dim DPP_DP As Decimal = 0
        Dim PPN_DP As Decimal = 0
        Dim DefPrecIDR As Integer = 0
        Dim FG_UANG_MUKA As Integer = 0
        Dim DPRef As String = ""
        Dim DummyDPP As Decimal = 0
        Dim DummyPPN As Decimal = 0
        Dim ItemCount As Integer = 0
        Dim Dummy1 As Decimal = 0

        Dim DTA5 As New DseLaporanVAT2025TableAdapters.DataTable5TableAdapter
        Dim Tbl5 As New DseLaporanVAT2025.DataTable5DataTable
        Dim DTA6 As New DseLaporanVAT2025TableAdapters.DataTable6TableAdapter
        Dim Tbl6 As New DseLaporanVAT2025.DataTable6DataTable
        Dim Row6 As DseLaporanVAT2025.DataTable6Row

        '#by ahdar | 19/06/20 | @terkait dpp yang di export dari laporan vat mengambil sub total + pengiriman
        Dim OngkosKirim As Decimal = 0
        Dim IsShipmentWithVAT As Boolean = False
        '#
        Try
            DTA5.Connection.ConnectionString = Core_Module.GetConstring
            DTA6.Connection.ConnectionString = Core_Module.GetConstring

            '~ab~
            'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
            '03/02/2020 
            Dim tblDummy As New DseLaporanVAT2025.Tbl_VATDataTable
            Dim rowDummy As DseLaporanVAT2025.Tbl_VATRow

            Dim ListExport As New List(Of Double)
            Dim ListExportStatus As New List(Of Double)
            ListExport.Clear()
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If row.Cells(Me.ColExport.Index).Value = Boolean.TrueString Then 'jika betul2 belum pernah di export
                    If ListExport.Contains(row.Cells(Me.ColPrimaryDocAcuan.Index).Value) = False Then
                        ListExport.Add(row.Cells(Me.ColPrimaryDocAcuan.Index).Value)
                    End If
                Else
                    If String.IsNullOrEmpty(row.Cells(Me.ColExportStatus.Index).Value) = False Then ' jika sudah pernah di export
                        If ListExportStatus.Contains(row.Cells(Me.ColPrimaryDocAcuan.Index).Value) = False Then
                            ListExportStatus.Add(row.Cells(Me.ColPrimaryDocAcuan.Index).Value)
                        End If
                    End If
                End If
            Next
            '~

            DefPrecIDR = Me.ObjCur.GetDecimal_DefaultCurrency
            'keluarkan datanya kembali
            'Me.btnShow.PerformClick()

            'cek berkaitan dgn no NPWP 
            For Each RowLaporan In DS.Tbl_VAT.Rows
                If ListExport.Contains(RowLaporan.primary_sales_invoice) = True Then
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
                End If
            Next

            If DummyIn <> "" Or DummyOut <> "" Then
                Dummy = "Please provide correct NPWP number for the following document:"
                If DummyIn <> "" Then
                    Dummy = Dummy + vbCrLf + DummyIn
                End If
                If DummyOut <> "" Then
                    Dummy = Dummy + vbCrLf + DummyOut
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
                    If ListExport.Contains(RowLaporan.primary_sales_invoice) = True Then
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
                    End If
                End With
            Next
            If DummyIn <> "" Or DummyOut <> "" Then
                Dummy = "Please provide correct tax invoice number for the following document:"
                If DummyIn <> "" Then
                    Dummy = Dummy + vbCrLf + DummyIn
                End If
                If DummyOut <> "" Then
                    Dummy = Dummy + vbCrLf + DummyOut
                End If

                MsgBox(Dummy, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '~ab~
            'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
            '03/02/2020
            If ListExport.Count = 0 Then
                MsgBox("Please check export minimum one item.", MsgBoxStyle.Information, "Information")
                Me.btnShow.PerformClick()
                Exit Sub
            End If

            tblDummy.Rows.Clear()
            tblDummy = DS.Tbl_VAT.Copy
            For i As Integer = tblDummy.Rows.Count - 1 To 0 Step -1
                rowDummy = tblDummy.Rows(i)
                If ListExport.Contains(rowDummy.primary_sales_invoice) = False Then
                    tblDummy.Rows.Remove(rowDummy)
                End If
            Next
            '~

            'If DS.Tbl_VAT.Rows.Count > 0 Then
            If tblDummy.Rows.Count > 0 Then
                If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor

                    'load data item
                    Me.DTA.Fill(Me.TBL, Me.dtpTo.Value, Me.dtpFrom.Value)
                    Me.DTAReturn.Fill(Me.TBLReturn, Me.dtpTo.Value, Me.dtpFrom.Value)
                    Me.KurangkanItemReturn()
                    Me.DTAItemDP.Fill(Me.TBLItemDP, Me.dtpTo.Value, Me.dtpFrom.Value)

                    DTA5.Fill(Tbl5, Me.dtpFrom.Value, Me.dtpTo.Value)
                    DTA6.Fill(Tbl6)
                    'For Each RowLaporan In DS.Tbl_VAT.Rows
                    '    With RowLaporan
                    For Each rowDummy In tblDummy.Rows
                        With rowDummy
                            KodeJenisTransaksiPajakDummy = .kode_jenis_transaksi_pajak
                            If KodeJenisTransaksiPajakDummy = "-1" Then
                                KodeJenisTransaksiPajakDummy = ""
                            Else
                                KodeJenisTransaksiPajakDummy = CInt(.kode_jenis_transaksi_pajak) + 1
                            End If
                            NoFakturDummy = .Tax_Invoice_No
                            NoFakturDummy = NoFakturDummy.Trim
                            NoFakturDummy = NoFakturDummy.Replace("-", "")
                            NoFakturDummy = NoFakturDummy.Replace(".", "")
                            NoNPWP = .NPWP
                            NoNPWP = NoNPWP.Trim
                            NoNPWP = NoNPWP.Replace("-", "")
                            NoNPWP = NoNPWP.Replace(".", "")
                            Dummy = .Alamat_Pajak_1 + " " + .Alamat_Pajak_2
                            If Dummy.Length > 255 Then
                                Dummy = Dummy.Substring(0, 255)
                            End If
                            .Cust_Supp_Name = .Cust_Supp_Name.Replace(";", "")
                            .Cust_Supp_Name = .Cust_Supp_Name.Replace("""", "")
                            .Invoice_No = .Invoice_No.Replace(";", "")
                            .Invoice_No = .Invoice_No.Replace("""", "")
                            .NPWP = .NPWP.Replace(";", "")
                            .NPWP = .NPWP.Replace("""", "")
                            .Retur = .Retur.Replace(";", "")
                            .Retur = .Retur.Replace("""", "")
                            .Tax_Address = .Tax_Address.Replace(";", "")
                            .Tax_Address = .Tax_Address.Replace("""", "")
                            .Tax_Invoice_No = .Tax_Invoice_No.Replace(";", "")
                            .Tax_Invoice_No = .Tax_Invoice_No.Replace("""", "")
                            .Type = .Type.Replace(";", "")
                            .Type = .Type.Replace("""", "")
                            .Nama_Jalan = .Nama_Jalan.Replace(";", "")
                            .Nama_Jalan = .Nama_Jalan.Replace("""", "")
                            .Blok = .Blok.Replace(";", "")
                            .Blok = .Blok.Replace("""", "")
                            .Nomor = .Nomor.Replace(";", "")
                            .Nomor = .Nomor.Replace("""", "")
                            .RT = .RT.Replace(";", "")
                            .RT = .RT.Replace("""", "")
                            .RW = .RW.Replace(";", "")
                            .RW = .RW.Replace("""", "")
                            .Kelurahan = .Kelurahan.Replace(";", "")
                            .Kelurahan = .Kelurahan.Replace("""", "")
                            .Kecamatan = .Kecamatan.Replace(";", "")
                            .Kecamatan = .Kecamatan.Replace("""", "")
                            .Kabupaten = .Kabupaten.Replace(";", "")
                            .Kabupaten = .Kabupaten.Replace("""", "")
                            .Propinsi = .Propinsi.Replace(";", "")
                            .Propinsi = .Propinsi.Replace("""", "")
                            .Kode_Pos = .Kode_Pos.Replace(";", "")
                            .Kode_Pos = .Kode_Pos.Replace("""", "")
                            .No_Telepon = .No_Telepon.Replace(";", "")
                            .No_Telepon = .No_Telepon.Replace("""", "")
                            '//TR : Tambahan baru Nik/Pas
                            .NikPass = .NikPass.Replace(";", "").Replace("""", "")
                            .SPPB_NO = .SPPB_NO.Replace(";", "").Replace("""", "")
                            Me.TBLItemDP.DefaultView.RowFilter = "[Primary Sales Invoice] = " + .primary_sales_invoice.ToString
                            DPP_DP = 0
                            PPN_DP = 0
                            If Me.TBLItemDP.DefaultView.Count = 0 Then
                                If .is_dp = 0 Then
                                    'tidak ada uang muka
                                    FG_UANG_MUKA = 0
                                ElseIf .is_dp = 1 Then
                                    'tidak ada uang muka
                                    FG_UANG_MUKA = 1
                                Else
                                    'tidak ada uang muka
                                    FG_UANG_MUKA = 0
                                End If
                            Else
                                'dianggap langsung pelunasan
                                FG_UANG_MUKA = 2
                            End If
                            For Each DV As DataRowView In Me.TBLItemDP.DefaultView
                                Me.rowItemDP = DV.Row
                                DPP_DP += Math.Round(Me.rowItemDP.Sub_Total * Me.rowItemDP.Tax_Exchange_Rate, DefPrecIDR, MidpointRounding.AwayFromZero)
                                PPN_DP += Math.Round(Me.rowItemDP.Tax_Amount * Me.rowItemDP.Tax_Exchange_Rate, DefPrecIDR, MidpointRounding.AwayFromZero)
                            Next

                            Me.TBL.DefaultView.RowFilter = "[Primary Invoice] = " + .primary_sales_invoice.ToString
                            If Me.TBL.DefaultView.Count > 0 Then
                                'header  
                                DummyDPP = ObjCD.RoundDown1(CDbl(.Purchase_Sales), 0)
                                DummyPPN = Math.Round(CDbl(.VAT), DefPrecIDR, MidpointRounding.AwayFromZero)
                                If FG_UANG_MUKA = 2 Then
                                    Dummy1 = DummyPPN
                                    DummyPPN = PPN_DP + Dummy1
                                    PPN_DP = Dummy1
                                    Dummy1 = DummyDPP
                                    DummyDPP = DPP_DP + Dummy1
                                    DPP_DP = Dummy1
                                End If

                                Dummy = Dummy.Replace(vbCrLf, "")
                                ValTxt = ValTxt + vbCrLf + "FK;" + NoFakturDummy.Substring(0, 2).ToCharArray + ";" + NoFakturDummy.Substring(2, 1) + ";" + NoFakturDummy.Substring(3) + ";" _
                                + CDate(._Date).Month.ToString + ";" + CDate(._Date).Year.ToString + ";" + Format(CDate(._Date), "dd/MM/yyyy") + ";" _
                                + NoNPWP + ";" + .Cust_Supp_Name + ";" + Dummy + ";" + Math.Round(DummyDPP, 0, MidpointRounding.AwayFromZero).ToString + ";" + _
                                Math.Round(DummyPPN, 0, MidpointRounding.AwayFromZero).ToString + ";0;" + _
                                KodeJenisTransaksiPajakDummy + ";" + FG_UANG_MUKA.ToString + ";" + CStr(Math.Round(CDbl(DPP_DP), 0, MidpointRounding.AwayFromZero)) + ";" + _
                                CStr(Math.Round(CDbl(PPN_DP), 0, MidpointRounding.AwayFromZero)) + ";0;" + _
                                Chr(34) & "Invoice No : " & .Invoice_No & " " & .NikPass '2024-12-19 IM tambah informasi Invoice No :
                                'Chr(34) & .Invoice_No & " " & .NikPass

                                Me.MasukanRefLainnya(ValTxt, Tbl5, Tbl6, .Invoice_No, ._Date, .Cust_Supp_Name)
                                ValTxt &= ";" & .SPPB_NO

                                '~ab~
                                'Remarks : penambahan batasan jika tidak ad ano npwp maka tidak keluar di export csv
                                '03/02/2020
                                If NoNPWP <> "000000000000000" Then
                                    'lawan transaksi
                                    ValTxt = ValTxt + vbCrLf + "LT" + ";" + NoNPWP + ";" + .Cust_Supp_Name + ";" + .Nama_Jalan + ";" + .Blok + ";" + .Nomor + ";" + _
                                    .RT + ";" + .RW + ";" + .Kecamatan + ";" + .Kelurahan + ";" + .Kabupaten + ";" + .Propinsi + ";" + .Kode_Pos + ";" + .No_Telepon
                                    'item
                                End If

                                '#by ahdar | 19/06/20 |
                                OngkosKirim = 0
                                IsShipmentWithVAT = False
                                '#
                                ItemCount = 0
                                For Each dv As DataRowView In Me.TBL.DefaultView
                                    ItemCount += 1
                                    Me.row = dv.Row
                                    If ItemCount = Me.TBL.DefaultView.Count Then
                                        '#by ahdar | 19/06/20 | @terkait dpp yang di export dari laporan vat mengambil sub total + pengiriman
                                        If Me.ObjInv.IsShipmentWithTax(Me.row.Primary_Invoice) = True Then
                                            OngkosKirim = Me.ObjInv.GetOngkosKirim(Me.row.Primary_Invoice)
                                            DummyDPP -= OngkosKirim
                                            DummyPPN = Math.Round(DummyDPP * 0.1, DefPrecIDR, MidpointRounding.AwayFromZero)
                                            IsShipmentWithVAT = True
                                        Else
                                            IsShipmentWithVAT = False
                                        End If
                                        '#
                                        Me.row.DPP_Value = DummyDPP
                                        Me.row.PPN_Value = DummyPPN
                                    Else
                                        Me.row.DPP_Value = Me.row.Sales_Amount
                                        Me.row.PPN_Value = Math.Round(Me.row.Sales_Amount * 0.1, DefPrecIDR, MidpointRounding.AwayFromZero)
                                        DummyDPP -= Me.row.DPP_Value
                                        DummyPPN -= Me.row.PPN_Value
                                    End If
                                Next


                                For Each dv As DataRowView In Me.TBL.DefaultView
                                    Me.row = dv.Row
                                    Me.row.Product_Code = Me.row.Product_Code.Replace("""", "")
                                    Me.row.Description = Me.row.Description.Replace("""", "")
                                    If Me.row.Tax_Exchange_Rate <> 1 Then
                                        If Not Me.ChkShowDiscount.Checked Then
                                            ValTxt = ValTxt + vbCrLf + "OF" + ";" + Me.row.Product_Code.Replace(";", ",") + ";" + Me.row.Description.Replace(";", ",") + ";" + _
                                            CStr(Me.ObjCD.RoundDown1(CDbl(Me.row.Nett_Price) * CDbl(Me.row.Tax_Exchange_Rate), 0)) + ";" + CStr(CDbl(Me.row.Total_Qty_Multi)) + ";" + CStr(CDbl(Me.row.DPP_Value)) + ";0;" + _
                                            CStr(CDbl(Me.row.DPP_Value)) + ";" + CStr(CDbl(Me.row.PPN_Value)) + ";0;0"
                                        Else
                                            Dim DiscAmount As Double = CDbl(Me.ObjCD.RoundDown1(CDbl(Me.row.Unit_Price) * CDbl(Me.row.Tax_Exchange_Rate), 0) * CDbl(Me.row.Total_Qty_Multi)) - CDbl(Me.row.DPP_Value)
                                            If DiscAmount < 0 Then DiscAmount = 0
                                            ValTxt = ValTxt + vbCrLf + "OF" + ";" + Me.row.Product_Code.Replace(";", ",") + ";" + Me.row.Description.Replace(";", ",") + ";" + _
                                            CStr(Me.ObjCD.RoundDown1(CDbl(Me.row.Unit_Price) * CDbl(Me.row.Tax_Exchange_Rate), 0)) + ";" + CStr(CDbl(Me.row.Total_Qty_Multi)) + _
                                            ";" + CStr(CDbl(Me.ObjCD.RoundDown1(CDbl(Me.row.Unit_Price) * CDbl(Me.row.Tax_Exchange_Rate), 0) * CDbl(Me.row.Total_Qty_Multi))) + ";" + CStr(DiscAmount) + ";" + _
                                            CStr(CDbl(Me.row.DPP_Value)) + ";" + CStr(CDbl(Me.row.PPN_Value)) + ";0;0"
                                        End If
                                    Else
                                        If Not Me.ChkShowDiscount.Checked Then
                                            ValTxt = ValTxt + vbCrLf + "OF" + ";" + Me.row.Product_Code.Replace(";", ",") + ";" + Me.row.Description.Replace(";", ",") + ";" + _
                                            CStr(CDbl(Me.row.Nett_Price)) + ";" + CStr(CDbl(Me.row.Total_Qty_Multi)) + ";" + CStr(CDbl(Me.row.Amount)) + ";0;" + _
                                            CStr(CDbl(Me.row.DPP_Value)) + ";" + CStr(CDbl(Me.row.PPN_Value)) + ";0;0"
                                        Else
                                            Dim DiscAmount As Double = Me.ObjCD.RoundDown1(CDbl(Me.row.Unit_Price) * CDbl(Me.row.Total_Qty_Multi), 0) - CDbl(Me.row.DPP_Value)
                                            If DiscAmount < 0 Then DiscAmount = 0
                                            ValTxt = ValTxt + vbCrLf + "OF" + ";" + Me.row.Product_Code.Replace(";", ",") + ";" + Me.row.Description.Replace(";", ",") + ";" + _
                                            CStr(CDbl(Me.row.Unit_Price)) + ";" + CStr(CDbl(Me.row.Total_Qty_Multi)) + ";" + CStr(Me.ObjCD.RoundDown1(CDbl(Me.row.Unit_Price) * CDbl(Me.row.Total_Qty_Multi), 0)) + ";" + CStr(DiscAmount) + ";" + _
                                            CStr(CDbl(Me.row.DPP_Value)) + ";" + CStr(CDbl(Me.row.PPN_Value)) + ";0;0"
                                        End If
                                    End If
                                Next

                                '#by ahdar | 19/06/20 |
                                If IsShipmentWithVAT = True And OngkosKirim > 0 Then
                                    Me.AddRow_Shipping(ValTxt, OngkosKirim, DefPrecIDR)
                                End If
                                '#
                            Else
                                'header
                                DummyDPP = ObjCD.RoundDown1(CDbl(.Purchase_Sales), 0)
                                DummyPPN = Math.Round(CDbl(.VAT), DefPrecIDR, MidpointRounding.AwayFromZero)
                                DPP_DP = DummyDPP
                                PPN_DP = DummyPPN
                                ValTxt = ValTxt + vbCrLf + "FK;" + NoFakturDummy.Substring(0, 2).ToCharArray + ";" + NoFakturDummy.Substring(2, 1) + ";" + NoFakturDummy.Substring(3) + ";" _
                                + CDate(._Date).Month.ToString + ";" + CDate(._Date).Year.ToString + ";" + Format(CDate(._Date), "dd/MM/yyyy") + ";" _
                                + NoNPWP + ";" + .Cust_Supp_Name + ";" + Dummy + ";" + Math.Round(DummyDPP, 0, MidpointRounding.AwayFromZero).ToString + ";" + _
                                Math.Round(DummyPPN, 0, MidpointRounding.AwayFromZero).ToString + ";0;" + _
                                KodeJenisTransaksiPajakDummy + ";" + FG_UANG_MUKA.ToString + ";" + Math.Round(DPP_DP, 0, MidpointRounding.AwayFromZero).ToString + ";" + _
                                CStr(Math.Round(CDbl(PPN_DP), 0, MidpointRounding.AwayFromZero)) + ";0;" + _
                                .Invoice_No & " " & .NikPass + ";" + .SPPB_NO
                                'lawan transaksi
                                ValTxt = ValTxt + vbCrLf + "LT" + ";" + NoNPWP + ";" + .Cust_Supp_Name + ";" + .Nama_Jalan + ";" + .Blok + ";" + .Nomor + ";" + _
                                .RT + ";" + .RW + ";" + .Kecamatan + ";" + .Kelurahan + ";" + .Kabupaten + ";" + .Propinsi + ";" + .Kode_Pos + ";" + .No_Telepon
                                'item

                                'item jika invoice dp
                                DPRef = Me.ObjInv.GetCustomerPONumber_ForInvoiceDP(.primary_sales_invoice)
                                If DPRef = "" Then
                                    DPRef = "down payment for SO: " + Me.ObjInv.GetSONumber_ForInvoiceDP(.primary_sales_invoice)
                                Else
                                    DPRef = "down payment for Cust. PO: " + DPRef
                                End If
                                ValTxt = ValTxt + vbCrLf + "OF" + ";DP;" + DPRef + ";" + _
                                CStr(DummyDPP) + ";1;" + CStr(DummyDPP) + ";0;" + _
                                CStr(DummyDPP) + ";" + CStr(DummyPPN) + ";0;0"
                            End If
                        End With
                    Next
                    My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, ValTxt, False)
                    MsgBox("Csv file created. Please be aware that characters used in import process (;,"") will be omitted.", MsgBoxStyle.Information, "Information")

                    For Each RowLaporan In DS.Tbl_VAT.Rows
                        '~ab~
                        'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
                        '03/02/2020
                        With RowLaporan
                            If ListExport.Contains(.primary_sales_invoice) = True Then
                                Me.UpdateExportStatus_Invoice(Me.UserCode & " on " & Format(Now, "dd MMM yyyy - HH.mm"), .primary_sales_invoice)
                            Else
                                If ListExportStatus.Contains(.primary_sales_invoice) = False Then
                                    Me.UpdateExportStatus_Invoice("-", .primary_sales_invoice)
                                End If
                            End If
                        End With
                    Next
                End If
            End If

            Me.btnShow.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddRow_Shipping(ByRef TxtVal As String, ByVal OngkosKirim As Decimal, ByVal DefPrecIDR As Integer)
        Dim Retval As String = TxtVal
        Dim Retval_Kol As String = ""
        Dim strArr() As String
        Dim strArr_Kol() As String
        Dim PPN As Decimal = 0
        Dim Split_String As String = ""
        Try
            strArr = Retval.Split(vbNewLine)
            Retval_Kol = strArr(strArr.Length - 1)
            If Retval.Contains(";") Then
                strArr_Kol = Retval_Kol.Split(";")
                Split_String = ";"
            Else
                strArr_Kol = Retval_Kol.Split(",")
                Split_String = ","
            End If

            'set kolom
            PPN = Math.Round(OngkosKirim * 0.1, DefPrecIDR, MidpointRounding.AwayFromZero)
            strArr_Kol(0) = strArr_Kol(0) & Split_String
            strArr_Kol(1) = "SHIPPING" & Split_String
            strArr_Kol(2) = "SHIPPING" & Split_String
            strArr_Kol(3) = CStr(CDbl(OngkosKirim)) & Split_String
            strArr_Kol(4) = 1 & Split_String
            strArr_Kol(5) = CStr(CDbl(OngkosKirim)) & Split_String
            strArr_Kol(6) = 0 & Split_String
            strArr_Kol(7) = CStr(CDbl(OngkosKirim)) & Split_String
            strArr_Kol(8) = CStr(CDbl(PPN)) & Split_String
            strArr_Kol(9) = strArr_Kol(9) & Split_String
            For i As Integer = 0 To strArr_Kol.Length - 1
                Retval &= strArr_Kol(i)
            Next
            TxtVal = Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub ExportPajakKeluaran()
    '    Dim ValTxt As String = "FK;KD_JENIS_TRANSAKSI;FG_PENGGANTI;NOMOR_FAKTUR;MASA_PAJAK;TAHUN_PAJAK;TANGGAL_FAKTUR;NPWP;NAMA;ALAMAT_LENGKAP;JUMLAH_DPP;JUMLAH_PPN;JUMLAH_PPNBM;ID_KETERANGAN_TAMBAHAN;FG_UANG_MUKA;UANG_MUKA_DPP;UANG_MUKA_PPN;UANG_MUKA_PPNBM;REFERENSI"
    '    ValTxt = ValTxt + vbCrLf + "LT;NPWP;NAMA;JALAN;BLOK;NOMOR;RT;RW;KECAMATAN;KELURAHAN;KABUPATEN;PROPINSI;KODE_POS;NO_TELEPON"
    '    ValTxt = ValTxt + vbCrLf + "OF;KODE_OBJEK;NAMA;HARGA_SATUAN;JUMLAH_BARANG;HARGA_TOTAL;DISKON;DPP;PPN;TARIF_PPNBM;PPNBM"

    '    Dim RowLaporan As DseLaporanVAT2025.Tbl_VATRow
    '    Dim DummyIn As String = ""
    '    Dim DummyOut As String = ""
    '    Dim Dummy As String = ""
    '    Dim KodeJenisTransaksiPajakDummy As String = ""
    '    Dim NoFakturDummy As String = ""
    '    Dim NoNPWP As String = ""
    '    Dim DPP_DP As Decimal = 0
    '    Dim PPN_DP As Decimal = 0
    '    Dim DefPrecIDR As Integer = 0
    '    Dim FG_UANG_MUKA As Integer = 0
    '    Dim DPRef As String = ""
    '    Dim DummyDPP As Decimal = 0
    '    Dim DummyPPN As Decimal = 0
    '    Dim ItemCount As Integer = 0
    '    Dim Dummy1 As Decimal = 0

    '    Try
    '        DefPrecIDR = Me.ObjCur.GetDecimal_DefaultCurrency
    '        'keluarkan datanya kembali
    '        Me.btnShow.PerformClick()

    '        'cek berkaitan dgn no NPWP 
    '        For Each RowLaporan In DS.Tbl_VAT.Rows
    '            If RowLaporan.VAT <> 0 Then
    '                With RowLaporan
    '                    If .Type = "VAT IN" Then
    '                        If RowLaporan.NPWP = "" Then
    '                            If DummyIn = "" Then
    '                                DummyIn = RowLaporan.Invoice_No
    '                            Else
    '                                DummyIn = DummyIn + ", " + RowLaporan.Invoice_No
    '                            End If
    '                        End If
    '                    Else
    '                        If RowLaporan.NPWP = "" Then
    '                            If DummyOut = "" Then
    '                                DummyOut = RowLaporan.Invoice_No
    '                            Else
    '                                DummyOut = DummyOut + ", " + RowLaporan.Invoice_No
    '                            End If
    '                        End If
    '                    End If
    '                End With
    '            End If
    '        Next
    '        If DummyIn <> "" Or DummyOut <> "" Then
    '            Dummy = "Please provide correct NPWP number for the following document:"
    '            If DummyIn <> "" Then
    '                Dummy = Dummy + vbCrLf + DummyIn
    '            End If
    '            If DummyOut <> "" Then
    '                Dummy = Dummy + vbCrLf + DummyOut
    '            End If

    '            MsgBox(Dummy, MsgBoxStyle.Information, "Information")
    '            Exit Sub
    '        End If

    '        'cek format tax invoice number yang berkaitan pembelian dan penjualan
    '        Dummy = ""
    '        DummyIn = ""
    '        DummyOut = ""
    '        For Each RowLaporan In DS.Tbl_VAT.Rows
    '            With RowLaporan
    '                If RowLaporan.VAT <> 0 Then
    '                    If RowLaporan.Type = "VAT IN" Then
    '                        If RowLaporan.Tax_Invoice_No = "" Or RowLaporan.Tax_Invoice_No.Length < 2 Then
    '                            If DummyIn = "" Then
    '                                DummyIn = RowLaporan.Invoice_No
    '                            Else
    '                                DummyIn = DummyIn + ", " + RowLaporan.Invoice_No
    '                            End If
    '                        End If
    '                    Else
    '                        If RowLaporan.Tax_Invoice_No = "" Or RowLaporan.Tax_Invoice_No.Length < 2 Then
    '                            If DummyOut = "" Then
    '                                DummyOut = RowLaporan.Invoice_No
    '                            Else
    '                                DummyOut = DummyOut + ", " + RowLaporan.Invoice_No
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            End With
    '        Next
    '        If DummyIn <> "" Or DummyOut <> "" Then
    '            Dummy = "Please provide correct tax invoice number for the following document:"
    '            If DummyIn <> "" Then
    '                Dummy = Dummy + vbCrLf + DummyIn
    '            End If
    '            If DummyOut <> "" Then
    '                Dummy = Dummy + vbCrLf + DummyOut
    '            End If

    '            MsgBox(Dummy, MsgBoxStyle.Information, "Information")
    '            Exit Sub
    '        End If

    '        If DS.Tbl_VAT.Rows.Count > 0 Then
    '            If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
    '                'load data item
    '                Me.DTA.Fill(Me.TBL, Me.dtpTo.Value, Me.dtpFrom.Value)
    '                Me.DTAReturn.Fill(Me.TBLReturn, Me.dtpTo.Value, Me.dtpFrom.Value)
    '                Me.KurangkanItemReturn()
    '                Me.DTAItemDP.Fill(Me.TBLItemDP, Me.dtpTo.Value, Me.dtpFrom.Value)

    '                For Each RowLaporan In DS.Tbl_VAT.Rows
    '                    With RowLaporan
    '                        KodeJenisTransaksiPajakDummy = .kode_jenis_transaksi_pajak
    '                        If KodeJenisTransaksiPajakDummy = "-1" Then
    '                            KodeJenisTransaksiPajakDummy = ""
    '                        Else
    '                            KodeJenisTransaksiPajakDummy = CInt(.kode_jenis_transaksi_pajak) + 1
    '                        End If
    '                        NoFakturDummy = RowLaporan.Tax_Invoice_No
    '                        NoFakturDummy = NoFakturDummy.Trim
    '                        NoFakturDummy = NoFakturDummy.Replace("-", "")
    '                        NoFakturDummy = NoFakturDummy.Replace(".", "")
    '                        NoNPWP = RowLaporan.NPWP
    '                        NoNPWP = NoNPWP.Trim
    '                        NoNPWP = NoNPWP.Replace("-", "")
    '                        NoNPWP = NoNPWP.Replace(".", "")
    '                        Dummy = .Alamat_Pajak_1 + " " + .Alamat_Pajak_2
    '                        If Dummy.Length > 255 Then
    '                            Dummy = Dummy.Substring(0, 255)
    '                        End If
    '                        .Cust_Supp_Name = .Cust_Supp_Name.Replace(";", "")
    '                        .Cust_Supp_Name = .Cust_Supp_Name.Replace("""", "")
    '                        .Invoice_No = .Invoice_No.Replace(";", "")
    '                        .Invoice_No = .Invoice_No.Replace("""", "")
    '                        .NPWP = .NPWP.Replace(";", "")
    '                        .NPWP = .NPWP.Replace("""", "")
    '                        .Retur = .Retur.Replace(";", "")
    '                        .Retur = .Retur.Replace("""", "")
    '                        .Tax_Address = .Tax_Address.Replace(";", "")
    '                        .Tax_Address = .Tax_Address.Replace("""", "")
    '                        .Tax_Invoice_No = .Tax_Invoice_No.Replace(";", "")
    '                        .Tax_Invoice_No = .Tax_Invoice_No.Replace("""", "")
    '                        .Type = .Type.Replace(";", "")
    '                        .Type = .Type.Replace("""", "")
    '                        .Nama_Jalan = .Nama_Jalan.Replace(";", "")
    '                        .Nama_Jalan = .Nama_Jalan.Replace("""", "")
    '                        .Blok = .Blok.Replace(";", "")
    '                        .Blok = .Blok.Replace("""", "")
    '                        .Nomor = .Nomor.Replace(";", "")
    '                        .Nomor = .Nomor.Replace("""", "")
    '                        .RT = .RT.Replace(";", "")
    '                        .RT = .RT.Replace("""", "")
    '                        .RW = .RW.Replace(";", "")
    '                        .RW = .RW.Replace("""", "")
    '                        .Kelurahan = .Kelurahan.Replace(";", "")
    '                        .Kelurahan = .Kelurahan.Replace("""", "")
    '                        .Kecamatan = .Kecamatan.Replace(";", "")
    '                        .Kecamatan = .Kecamatan.Replace("""", "")
    '                        .Kabupaten = .Kabupaten.Replace(";", "")
    '                        .Kabupaten = .Kabupaten.Replace("""", "")
    '                        .Propinsi = .Propinsi.Replace(";", "")
    '                        .Propinsi = .Propinsi.Replace("""", "")
    '                        .Kode_Pos = .Kode_Pos.Replace(";", "")
    '                        .Kode_Pos = .Kode_Pos.Replace("""", "")
    '                        .No_Telepon = .No_Telepon.Replace(";", "")
    '                        .No_Telepon = .No_Telepon.Replace("""", "")

    '                        Me.TBLItemDP.DefaultView.RowFilter = "[Primary Sales Invoice] = " + RowLaporan.primary_sales_invoice.ToString
    '                        DPP_DP = 0
    '                        PPN_DP = 0
    '                        If Me.TBLItemDP.DefaultView.Count = 0 Then
    '                            'tidak ada uang muka
    '                            FG_UANG_MUKA = 0
    '                        Else
    '                            'dianggap langsung pelunasan
    '                            FG_UANG_MUKA = 2
    '                        End If
    '                        For Each DV As DataRowView In Me.TBLItemDP.DefaultView
    '                            Me.rowItemDP = DV.Row
    '                            DPP_DP += Math.Round(Me.rowItemDP.Sub_Total * Me.rowItemDP.Tax_Exchange_Rate, DefPrecIDR, MidpointRounding.AwayFromZero)
    '                            PPN_DP += Math.Round(Me.rowItemDP.Tax_Amount * Me.rowItemDP.Tax_Exchange_Rate, DefPrecIDR, MidpointRounding.AwayFromZero)
    '                        Next
    '                        'header
    '                        DummyDPP = Math.Round(CDbl(.Purchase_Sales), DefPrecIDR, MidpointRounding.AwayFromZero)
    '                        DummyPPN = Math.Round(CDbl(.VAT), DefPrecIDR, MidpointRounding.AwayFromZero)
    '                        ValTxt = ValTxt + vbCrLf + "FK;" + NoFakturDummy.Substring(0, 2).ToCharArray + ";"+NoFakturDummy.Substring( 2,1)+";" + NoFakturDummy.Substring(3) + ";" _
    '                        + CDate(._Date).Month.ToString + ";" + CDate(._Date).Year.ToString + ";" + Format(CDate(._Date), "dd/MM/yyyy") + ";" _
    '                        + NoNPWP + ";" + .Cust_Supp_Name + ";" + Dummy + ";" + CStr(DummyDPP) + ";" + CStr(DummyPPN) + ";0;" + _
    '                        KodeJenisTransaksiPajakDummy + ";" + FG_UANG_MUKA.ToString + ";" + CStr(Math.Round(CDbl(DPP_DP), DefPrecIDR, MidpointRounding.AwayFromZero)) + ";" + _
    '                        CStr(Math.Round(CDbl(PPN_DP), DefPrecIDR, MidpointRounding.AwayFromZero)) + ";0;" + _
    '                        .Invoice_No
    '                        'lawan transaksi
    '                        ValTxt = ValTxt + vbCrLf + "LT" + ";" + NoNPWP + ";" + .Cust_Supp_Name + ";" + .Nama_Jalan + ";" + .Blok + ";" + .Nomor + ";" + _
    '                        .RT + ";" + .RW + ";" + .Kecamatan + ";" + .Kelurahan + ";" + .Kabupaten + ";" + .Propinsi + ";" + .Kode_Pos + ";" + .No_Telepon
    '                        'item
    '                        Me.TBL.DefaultView.RowFilter = "[Primary Invoice] = " + .primary_sales_invoice.ToString
    '                        If Me.TBL.DefaultView.Count > 0 Then
    '                            ItemCount = 0
    '                            For Each dv As DataRowView In Me.TBL.DefaultView
    '                                ItemCount += 1
    '                                Me.row = dv.Row
    '                                If ItemCount = Me.TBL.DefaultView.Count Then
    '                                    Me.row.DPP_Value = DummyDPP
    '                                    Me.row.PPN_Value = DummyPPN
    '                                Else
    '                                    If Me.row.Sales_Amount < DummyDPP Then
    '                                        Me.row.DPP_Value = Me.row.Sales_Amount
    '                                    Else
    '                                        Me.row.DPP_Value = DummyPPN
    '                                    End If
    '                                    Dummy1 = Math.Round(Me.row.Sales_Amount * 0.1, DefPrecIDR, MidpointRounding.AwayFromZero)
    '                                    If Dummy1 < DummyPPN Then
    '                                        Me.row.PPN_Value = Dummy1
    '                                    Else
    '                                        Me.row.PPN_Value = DummyPPN
    '                                    End If
    '                                    DummyDPP -= Me.row.DPP_Value
    '                                    DummyPPN -= Me.row.PPN_Value
    '                                End If
    '                            Next
    '                            For Each dv As DataRowView In Me.TBL.DefaultView
    '                                Me.row = dv.Row
    '                                Me.row.Product_Code = Me.row.Product_Code.Replace("""", "")
    '                                Me.row.Description = Me.row.Description.Replace("""", "")
    '                                If Me.row.Tax_Exchange_Rate <> 1 Then
    '                                    ValTxt = ValTxt + vbCrLf + "OF" + ";" + Me.row.Product_Code.Replace(";", ",") + ";" + Me.row.Description.Replace(";", ",") + ";" + _
    '                                    CStr(Me.ObjCD.RoundDown1(CDbl(Me.row.Nett_Price) * CDbl(Me.row.Tax_Exchange_Rate), 0)) + ";" + CStr(CDbl(Me.row.Total_Qty)) + ";" + CStr(CDbl(Me.row.DPP_Value)) + ";0;" + _
    '                                    CStr(CDbl(Me.row.DPP_Value)) + ";" + CStr(CDbl(Me.row.PPN_Value)) + ";0;0"
    '                                Else
    '                                    ValTxt = ValTxt + vbCrLf + "OF" + ";" + Me.row.Product_Code.Replace(";", ",") + ";" + Me.row.Description.Replace(";", ",") + ";" + _
    '                                    CStr(CDbl(Me.row.Nett_Price)) + ";" + CStr(CDbl(Me.row.Total_Qty)) + ";" + CStr(CDbl(Me.row.Amount)) + ";0;" + _
    '                                    CStr(CDbl(Me.row.DPP_Value)) + ";" + CStr(CDbl(Me.row.PPN_Value)) + ";0;0"
    '                                End If
    '                            Next
    '                        Else
    '                            'item jika invoice dp
    '                            DPRef = Me.ObjInv.GetCustomerPONumber_ForInvoiceDP(.primary_sales_invoice)
    '                            If DPRef = "" Then
    '                                DPRef = "down payment for SO: " + Me.ObjInv.GetSONumber_ForInvoiceDP(.primary_sales_invoice)
    '                            Else
    '                                DPRef = "down payment for Cust. PO: " + DPRef
    '                            End If
    '                            ValTxt = ValTxt + vbCrLf + "OF" + ";DP;" + DPRef + ";" + _
    '                            CStr(DummyDPP) + ";1;" + CStr(DummyDPP) + ";0;" + _
    '                            CStr(DummyDPP) + ";" + CStr(DummyPPN) + ";0;0"
    '                        End If
    '                    End With
    '                Next
    '                My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, ValTxt, False)
    '                MsgBox("Csv file created. Please be aware that characters used in import process (;,"") will be omitted.", MsgBoxStyle.Information, "Information")
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub btnShow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim TypeInOut As Integer = 1
        Dim SalerOrPr As Decimal = 0
        Dim VAT As Decimal = 0

        Dim RowIn As DseLaporanVAT2025.DataTable1Row
        Dim RowOut As DseLaporanVAT2025.DataTable2Row

        Dim RowDgv As DataGridViewRow
        Dim PrimaryGRNMin As Double = Double.MaxValue
        Dim PrimaryGRNMax As Double = Double.MinValue
        Dim TipeSalesInv As Sales_Module.ClsTransaksiSalesInvoice.EnumTipeInvoice = Sales_Module.ClsTransaksiSalesInvoice.EnumTipeInvoice.InvoiceDP
        Me.ColNoNotaRetur.Visible = False
        Dim AdaNPWP As Boolean = False
        Dim AdaNIK As Boolean = False
        Dim AdaPASS As Boolean = False
        Dim AdaOtherID As Boolean = False '2025 Core Tax Tahap 2.1
        Dim EndLicenseDate As Date = #1/1/1950#
        'Update NPWP 16 Digit
        Dim objNPWP As New Core_Module.ClsTrans_SettingNPWPNo

        '2025 Core Tax Tahap 2.1
        Dim TBLOut_SalesPelunasan As New DseLaporanVAT2025.InvoiceDataTable
        Dim rowOut_pelunasan As DseLaporanVAT2025.InvoiceRow
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

            '2025-Coretax
            Dim expirationDate As Date = New Date(2024, 12, 31)
            If Me.dtpFrom.Enabled = True Then
                If Me.dtpFrom.Value.Date <= expirationDate Then
                    MsgBox("This feature is only available after 1 January 2025.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo.Enabled = True Then
                If Me.dtpTo.Value.Date <= expirationDate Then
                    MsgBox("This feature is only available after 1 January 2025.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            'end 2025-Coretax


            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.ColInvoiceNumber.HeaderText = "Invoice Number"
            Me.ColTaxInvoiceNumber.HeaderText = "Tax Invoice Number"
            Me.ColPurchaseSales.HeaderText = "Purchase/Sales"
            Me.ColSuppCust.HeaderText = "Supplier/Customer"

            Me.DS.Tbl_VAT.Rows.Clear()
            crv1.ReportSource = Nothing
            crv1.Show()

            Me.dgvItem.Rows.Clear()
            TypeInOut = Me.ComboBox1.SelectedIndex

            If TypeInOut = EnumVATType.VAT_OUT_INVOICE Then
                Me.ColPurchaseSales.HeaderText = "Sales"
                Me.ColSuppCust.HeaderText = "Customer"
                'Me.DTAOut_Sales.Fill(TBLOut_Sales, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
                '2025 Core Tax Tahap 2.1 hanya invoice saja
                Me.DTAOut_Sales.FillBy_PeriodBy_TipeInvoice(TBLOut_Sales, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, _
                Sales_Module.ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice)
                Me.DTAOut_Sales.FillBy_PeriodWith_DPItem(TBLOut_SalesPelunasan, Me.dtpFrom.Value.Date, Me.dtpTo.Value, _
                Sales_Module.ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice)

                '#by ahdar | 20/06/20 | @terkait dpp yang di export dari laporan vat memperhitungkan ongkos kirim jika pakai
                For Each Me.rowOut_Sales In TBLOut_Sales.Rows
                    If Me.ObjInv.IsShipmentWithTax(Me.rowOut_Sales.Primary) = False Then
                        If Me.ObjInv.GetOngkosKirim(Me.rowOut_Sales.Primary) > 0 Then
                            Me.rowOut_Sales.Sales = Me.rowOut_Sales.Sales - Me.ObjInv.GetOngkosKirim(Me.rowOut_Sales.Primary)
                        End If
                    End If
                Next
                '#

                '2025 Core Tax Tahap 2.1
                'Cek NPWP Customer Invalid
                Dim DummyNPWP As String = ""
                Dim dt As New DataTable()
                dt.Columns.Add("ColInvoiceNo")
                dt.Columns.Add("ColInvoiceDate")
                dt.Columns.Add("ColCustomerName")
                dt.Columns.Add("ColDocType")
                dt.Columns.Add("ColDocNo")
                dt.Columns.Add("colPrimaryCustomer")
                dt.Columns.Add("colPrimaryInvoice")
                For Each Me.rowOut_Sales In TBLOut_Sales.Rows
                    If String.IsNullOrEmpty(rowOut_Sales.New_NPWP) = False Then
                        DummyNPWP = rowOut_Sales.New_NPWP.Replace(".", "").Replace("-", "").Replace(" ", "")
                    Else
                        DummyNPWP = rowOut_Sales.NPWP.Replace(".", "").Replace("-", "").Replace(" ", "")
                    End If

                    'cek lenght npwp
                    If DummyNPWP.Length < 16 Then
                        dt.Rows.Add()
                        dt.Rows(dt.Rows.Count - 1)("colPrimaryCustomer") = rowOut_Sales.Primary_Customer
                        dt.Rows(dt.Rows.Count - 1)("colPrimaryInvoice") = rowOut_Sales.Primary
                        dt.Rows(dt.Rows.Count - 1)("ColInvoiceNo") = rowOut_Sales.Invoice_no
                        dt.Rows(dt.Rows.Count - 1)("ColInvoiceDate") = rowOut_Sales.Tanggal_dari
                        dt.Rows(dt.Rows.Count - 1)("ColCustomerName") = rowOut_Sales.Cust
                        If DummyNPWP <> "000000000000000" Then
                            dt.Rows(dt.Rows.Count - 1)("ColDocType") = "NPWP"
                            If String.IsNullOrEmpty(rowOut_Sales.New_NPWP) = False Then
                                dt.Rows(dt.Rows.Count - 1)("ColDocNo") = rowOut_Sales.New_NPWP
                            Else
                                dt.Rows(dt.Rows.Count - 1)("ColDocNo") = rowOut_Sales.NPWP
                            End If
                        Else
                            If rowOut_Sales.Pilih_NIKPASS = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByNIK Then
                                dt.Rows(dt.Rows.Count - 1)("ColDocType") = "NIK"
                            ElseIf rowOut_Sales.Pilih_NIKPASS = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByPASSPORT Then
                                dt.Rows(dt.Rows.Count - 1)("ColDocType") = "PASSPORT"
                            Else
                                dt.Rows(dt.Rows.Count - 1)("ColDocType") = "OTHER ID"
                            End If
                            dt.Rows(dt.Rows.Count - 1)("ColDocNo") = Me.ObjCustomer.GetNoNikPass2(rowOut_Sales.Primary_Customer, rowOut_Sales.Pilih_NIKPASS)
                        End If
                    End If
                Next
                If dt.Rows.Count > 0 Then
                    Dim frmInvalid As New frmLaporanVAT2025_InvalidNPWP
                    frmInvalid.UserCode = Me.UserCode
                    frmInvalid.BackColorAll = Me.BackColorAll
                    frmInvalid._dt = dt
                    MsgBox("There is " & dt.Rows.Count & " data with npwp 15 digits")
                    frmInvalid.ShowDialog()
                    Exit Sub
                End If
                'end Cek NPWP Customer Invalid

                '2025 Core Tax Tahap 2.1
                'Cek transaksi Pelunasan, jika ada pada transaksi, maka hapus
                If TBLOut_SalesPelunasan.Rows.Count > 0 Then
                    Dim dv As DataRowView
                    Dim i As Integer = 0
                    For Each rowOut_pelunasan In TBLOut_SalesPelunasan
                        Me.TBLOut_Sales.DefaultView.RowFilter = "[Primary]=" & rowOut_pelunasan.Primary
                        If Me.TBLOut_Sales.DefaultView.Count > 0 Then
                            For i = Me.TBLOut_Sales.DefaultView.Count - 1 To 0 Step -1
                                dv = Me.TBLOut_Sales.DefaultView(i)
                                rowOut_Sales = dv.Row
                                Me.TBLOut_Sales.RemoveInvoiceRow(rowOut_Sales)
                            Next
                        End If
                    Next
                    Dim ij As New FrmLaporanVAT2025_ListInvoicePelunasan
                    ij.UserCode = Me.UserCode
                    ij.BackColorAll = Me.BackColorAll
                    'ij.TBLOut_SalesPelunasan = TBLOut_SalesPelunasan
                    ij.DateFrom = Me.dtpFrom.Value
                    ij.DateTo = Me.dtpTo.Value
                    ij.ShowDialog()
                End If
                '

                For Each rowOut_Sales In Me.TBLOut_Sales
                    With rowOut_Sales
                        'If .VAT_Dipungut = 0 Then
                        '    SalerOrPr = 0
                        '    VAT = 0
                        'Else : .VAT_Dipungut = 1
                        SalerOrPr = .Sales
                        VAT = .VAT
                        'End If
                        Me.dgvItem.Rows.Add()
                        RowDgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)

                        '~ab~
                        'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
                        '14/01/2020
                        If String.IsNullOrEmpty(.Export_Status) = False Then
                            RowDgv.Cells(Me.ColExportStatus.Index).Value = .Export_Status
                        Else
                            RowDgv.Cells(Me.ColExportStatus.Index).Value = ""
                        End If
                        '~

                        RowDgv.Cells(Me.ColTipeTransaksiString.Index).Value = .Out
                        RowDgv.Cells(Me.ColSuppCust.Index).Value = .Cust
                        RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                        RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no
                        '2025 Core Tax Tahap 2.2 
                        'RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .No_Faktur_Pajak
                        If String.IsNullOrEmpty(.Tax_Type) And .Tax_Type <> -1 Then
                            RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .No_Faktur_Pajak
                        Else
                            If .Tax_Revised = 0 Then
                                RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = Me.getFakturString(.Tax_Type) & "00" & .No_Faktur
                            Else
                                RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = Me.getFakturString(.Tax_Type) & "01" & .No_Faktur
                            End If
                        End If
                        RowDgv.Cells(Me.ColDate.Index).Value = Format(.Tanggal_dari.Date, "dd MMM yyyy")
                        RowDgv.Cells(Me.ColPurchaseSales.Index).Value = Me.ObjCD.DigitGrouping(SalerOrPr.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(VAT.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColPrimaryDocAcuan.Index).Value = .Primary
                        RowDgv.Cells(Me.ColEnumType.Index).Value = .Retur
                        RowDgv.Cells(Me.ColEnumInOut.Index).Value = FrmLaporanVAT.EnumInOut.VATOut
                        RowDgv.Cells(Me.ColTaxAddress.Index).Value = ""
                        RowDgv.Cells(Me.ColTanggalTransaksiRef.Index).Value = Format(#1/1/1950#, "dd MMM yyyy")
                        RowDgv.Cells(Me.ColNoNotaRetur.Index).Value = ""
                        RowDgv.Cells(Me.ColNoFakturPajakRef.Index).Value = ""
                        RowDgv.Cells(Me.ColNamaJalan.Index).Value = .Nama_Jalan
                        RowDgv.Cells(Me.ColBlok.Index).Value = .Blok
                        RowDgv.Cells(Me.ColNomor.Index).Value = .Nomor
                        RowDgv.Cells(Me.ColRT.Index).Value = .RT
                        RowDgv.Cells(Me.ColRW.Index).Value = .RW
                        RowDgv.Cells(Me.ColKelurahan.Index).Value = .Kelurahan
                        RowDgv.Cells(Me.ColKecamatan.Index).Value = .Kecamatan
                        RowDgv.Cells(Me.ColKabupaten.Index).Value = .Kabupaten
                        RowDgv.Cells(Me.ColPropinsi.Index).Value = .Propinsi
                        RowDgv.Cells(Me.ColKodePos.Index).Value = .Kode_Pos
                        RowDgv.Cells(Me.ColTelepon.Index).Value = .Telepon
                        RowDgv.Cells(Me.ColAlamatPajak1.Index).Value = .Alamat_Pajak_1
                        RowDgv.Cells(Me.ColAlamatPajak2.Index).Value = .Alamat_Pajak_2
                        RowDgv.Cells(Me.ColKodeJenisTransaksi.Index).Value = .kode_jenis_transaksi_pajak

                        'update NPWP 16 Digit.
                        If String.IsNullOrEmpty(.New_NPWP) = False Then
                            RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .New_NPWP
                        Else
                            If CDate(RowDgv.Cells(Me.ColDate.Index).Value).Date >= objNPWP.Tglberlaku Then
                                RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                            Else
                                RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP_Old 'sumber dari master customer 
                            End If
                        End If

                        '//TR : Tambahan baru Nik/Pas
                        'update NPWP 16 Digit.
                        If .Pilih_NIKPASS = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByNIK Then
                            RowDgv.Cells(Me.ColNik.Index).Value = Me.ObjCustomer.GetNoNikPass2(.Primary_Customer, .Pilih_NIKPASS)
                        ElseIf .Pilih_NIKPASS = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByPASSPORT Then
                            RowDgv.Cells(Me.ColPassport.Index).Value = Me.ObjCustomer.GetNoNikPass2(.Primary_Customer, .Pilih_NIKPASS)
                        Else '2025 Core Tax Tahap 2.1
                            RowDgv.Cells(Me.ColOtherID.Index).Value = Me.ObjCustomer.GetNoNikPass2(.Primary_Customer, .Pilih_NIKPASS)
                        End If

                        ''//TR : Tambahan baru Nik/Pas
                        'If .Pilih_NIKPASS = OptionNIKPASS.EnumOptionNikPass.ByNIK Then
                        '    'RowDgv.Cells(Me.ColNik.Index).Value = Me.ObjCustomer.GetNoNikPass2(.Primary_Customer, .Pilih_NIKPASS)
                        '    RowDgv.Cells(Me.ColNik.Index).Value = Me.ObjInv.GetNoNikPassbyInvoice(.Primary)
                        'Else
                        '    'RowDgv.Cells(Me.ColPassport.Index).Value = Me.ObjCustomer.GetNoNikPass2(.Primary_Customer, .Pilih_NIKPASS)
                        '    RowDgv.Cells(Me.ColPassport.Index).Value = Me.ObjInv.GetNoNikPassbyInvoice(.Primary)
                        'End If

                        '2025 Core Tax Tahap 2.1
                        RowDgv.Cells(Me.colNPWPType.Index).Value = .NPWP_Type
                        If RowDgv.Cells(Me.ColNPWPNumber.Index).Value <> "00.000.000.0-000.000" And RowDgv.Cells(Me.ColNPWPNumber.Index).Value <> "-" _
                        And RowDgv.Cells(Me.ColNPWPNumber.Index).Value <> "" And RowDgv.Cells(Me.ColNPWPNumber.Index).Value <> "0000 0000 0000 0000" _
                        And RowDgv.Cells(Me.ColNPWPNumber.Index).Value <> "0000000000000000" And .NPWP_Type <> Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP.Individual Then
                            AdaNPWP = True
                        End If
                        If RowDgv.Cells(Me.ColNik.Index).Value <> "-" And RowDgv.Cells(Me.ColNik.Index).Value <> "" Then
                            AdaNIK = True
                        End If
                        If RowDgv.Cells(Me.ColPassport.Index).Value <> "-" And RowDgv.Cells(Me.ColPassport.Index).Value <> "" Then
                            AdaPASS = True
                        End If
                        '2025 Core Tax Tahap 2.1
                        If RowDgv.Cells(Me.ColOtherID.Index).Value <> "" Then
                            AdaOtherID = True
                        End If

                        If .Tipe_Invoice = TipeSalesInv Then
                            'dp
                            RowDgv.Cells(Me.ColIsDP.Index).Value = 1
                        Else
                            'invoice biasa/pelunasan
                            RowDgv.Cells(Me.ColIsDP.Index).Value = 0
                        End If

                        RowDgv.Cells(Me.ColSPPBNo.Index).Value = .SPPB_No
                        RowDgv.Cells(Me.colEmail.Index).Value = .Email '2025-Coretax
                        RowDgv.Cells(Me.ColIDTKU.Index).Value = .TKU_ID '2025-Coretax

                        '2025 Core Tax Tahap 2.2
                        RowDgv.Cells(Me.colAdditionalInfo.Index).Value = .Additional_Info
                        RowDgv.Cells(Me.ColFacilityStamp.Index).Value = .Cap_Fasilitas
                        RowDgv.Cells(Me.colSPPBDate.Index).Value = .SPPB_Date
                        RowDgv.Cells(Me.colTaxRevised.Index).Value = .Tax_Revised
                        RowDgv.Cells(Me.colTaxType.Index).Value = .Tax_Type
                    End With
                Next
                '//TR : Tambahan baru Nik/Pas
                If AdaNPWP = False Then
                    Me.ColNPWPNumber.Visible = False
                Else
                    Me.ColNPWPNumber.Visible = True
                End If

                If AdaNIK = False Then
                    Me.ColNik.Visible = False
                Else
                    Me.ColNik.Visible = True
                End If
                If AdaPASS = False Then
                    Me.ColPassport.Visible = False
                Else
                    Me.ColPassport.Visible = True
                End If
                '2025 Core Tax Tahap 2.1
                If AdaOtherID = False Then
                    Me.ColOtherID.Visible = False
                Else
                    Me.ColOtherID.Visible = True
                End If
                Me.ColSPPBNo.Visible = True

                Me.SalinCrv()
            ElseIf TypeInOut = EnumVATType.VAT_OUT_INVOICEDP Then
                Me.ColPurchaseSales.HeaderText = "Sales"
                Me.ColSuppCust.HeaderText = "Customer"
                'Me.DTAOut_Sales.Fill(TBLOut_Sales, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
                '2025 Core Tax Tahap 2.1 hanya invoice saja
                Me.DTAOut_Sales.FillBy_PeriodBy_TipeInvoice(TBLOut_Sales, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, _
                Sales_Module.ClsTransaksiSalesInvoice.EnumTipeInvoice.InvoiceDP)



                '#by ahdar | 20/06/20 | @terkait dpp yang di export dari laporan vat memperhitungkan ongkos kirim jika pakai
                For Each Me.rowOut_Sales In TBLOut_Sales.Rows
                    If Me.ObjInv.IsShipmentWithTax(Me.rowOut_Sales.Primary) = False Then
                        If Me.ObjInv.GetOngkosKirim(Me.rowOut_Sales.Primary) > 0 Then
                            Me.rowOut_Sales.Sales = Me.rowOut_Sales.Sales - Me.ObjInv.GetOngkosKirim(Me.rowOut_Sales.Primary)
                        End If
                    End If
                Next
                '#

                '2025 Core Tax Tahap 2.1
                'Cek NPWP Customer Invalid
                Dim DummyNPWP As String = ""
                Dim dt As New DataTable()
                dt.Columns.Add("ColInvoiceNo")
                dt.Columns.Add("ColInvoiceDate")
                dt.Columns.Add("ColCustomerName")
                dt.Columns.Add("ColDocType")
                dt.Columns.Add("ColDocNo")
                dt.Columns.Add("colPrimaryCustomer")
                dt.Columns.Add("colPrimaryInvoice")
                For Each Me.rowOut_Sales In TBLOut_Sales.Rows
                    If String.IsNullOrEmpty(rowOut_Sales.New_NPWP) = False Then
                        DummyNPWP = rowOut_Sales.New_NPWP.Replace(".", "").Replace("-", "").Replace(" ", "")
                    Else
                        DummyNPWP = rowOut_Sales.NPWP.Replace(".", "").Replace("-", "").Replace(" ", "")
                    End If

                    'cek lenght npwp
                    If DummyNPWP.Length < 16 Then
                        dt.Rows.Add()
                        dt.Rows(dt.Rows.Count - 1)("colPrimaryCustomer") = rowOut_Sales.Primary_Customer
                        dt.Rows(dt.Rows.Count - 1)("colPrimaryInvoice") = rowOut_Sales.Primary
                        dt.Rows(dt.Rows.Count - 1)("ColInvoiceNo") = rowOut_Sales.Invoice_no
                        dt.Rows(dt.Rows.Count - 1)("ColInvoiceDate") = rowOut_Sales.Tanggal_dari
                        dt.Rows(dt.Rows.Count - 1)("ColCustomerName") = rowOut_Sales.Cust
                        If DummyNPWP <> "000000000000000" Then
                            dt.Rows(dt.Rows.Count - 1)("ColDocType") = "NPWP"
                            If String.IsNullOrEmpty(rowOut_Sales.New_NPWP) = False Then
                                dt.Rows(dt.Rows.Count - 1)("ColDocNo") = rowOut_Sales.New_NPWP
                            Else
                                dt.Rows(dt.Rows.Count - 1)("ColDocNo") = rowOut_Sales.NPWP
                            End If
                        Else
                            If rowOut_Sales.Pilih_NIKPASS = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByNIK Then
                                dt.Rows(dt.Rows.Count - 1)("ColDocType") = "NIK"
                            ElseIf rowOut_Sales.Pilih_NIKPASS = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByPASSPORT Then
                                dt.Rows(dt.Rows.Count - 1)("ColDocType") = "PASSPORT"
                            Else
                                dt.Rows(dt.Rows.Count - 1)("ColDocType") = "OTHER ID"
                            End If
                            dt.Rows(dt.Rows.Count - 1)("ColDocNo") = Me.ObjCustomer.GetNoNikPass2(rowOut_Sales.Primary_Customer, rowOut_Sales.Pilih_NIKPASS)
                        End If
                    End If
                Next
                If dt.Rows.Count > 0 Then
                    Dim frmInvalid As New frmLaporanVAT2025_InvalidNPWP
                    frmInvalid.UserCode = Me.UserCode
                    frmInvalid.BackColorAll = Me.BackColorAll
                    frmInvalid._dt = dt
                    MsgBox("Ada " & dt.Rows.Count & " data dengan npwp 15 digit")
                    frmInvalid.ShowDialog()
                    Exit Sub
                End If
                'end Cek NPWP Customer Invalid

                For Each rowOut_Sales In Me.TBLOut_Sales
                    With rowOut_Sales
                        'If .VAT_Dipungut = 0 Then
                        '    SalerOrPr = 0
                        '    VAT = 0
                        'Else : .VAT_Dipungut = 1
                        SalerOrPr = .Sales
                        VAT = .VAT
                        'End If
                        Me.dgvItem.Rows.Add()
                        RowDgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)

                        '~ab~
                        'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
                        '14/01/2020
                        If String.IsNullOrEmpty(.Export_Status) = False Then
                            RowDgv.Cells(Me.ColExportStatus.Index).Value = .Export_Status
                        Else
                            RowDgv.Cells(Me.ColExportStatus.Index).Value = ""
                        End If
                        '~

                        RowDgv.Cells(Me.ColTipeTransaksiString.Index).Value = .Out
                        RowDgv.Cells(Me.ColSuppCust.Index).Value = .Cust
                        RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                        RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no
                        '2025 Core Tax Tahap 2.2 
                        'RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .No_Faktur_Pajak
                        If String.IsNullOrEmpty(.Tax_Type) And .Tax_Type <> -1 Then
                            RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .No_Faktur_Pajak
                        Else
                            If .Tax_Revised = 0 Then
                                RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = Me.getFakturString(.Tax_Type) & "00" & .No_Faktur
                            Else
                                RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = Me.getFakturString(.Tax_Type) & "01" & .No_Faktur
                            End If
                        End If
                        RowDgv.Cells(Me.ColDate.Index).Value = Format(.Tanggal_dari.Date, "dd MMM yyyy")
                        RowDgv.Cells(Me.ColPurchaseSales.Index).Value = Me.ObjCD.DigitGrouping(SalerOrPr.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(VAT.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColPrimaryDocAcuan.Index).Value = .Primary
                        RowDgv.Cells(Me.ColEnumType.Index).Value = .Retur
                        RowDgv.Cells(Me.ColEnumInOut.Index).Value = FrmLaporanVAT.EnumInOut.VATOut
                        RowDgv.Cells(Me.ColTaxAddress.Index).Value = ""
                        RowDgv.Cells(Me.ColTanggalTransaksiRef.Index).Value = Format(#1/1/1950#, "dd MMM yyyy")
                        RowDgv.Cells(Me.ColNoNotaRetur.Index).Value = ""
                        RowDgv.Cells(Me.ColNoFakturPajakRef.Index).Value = ""
                        RowDgv.Cells(Me.ColNamaJalan.Index).Value = .Nama_Jalan
                        RowDgv.Cells(Me.ColBlok.Index).Value = .Blok
                        RowDgv.Cells(Me.ColNomor.Index).Value = .Nomor
                        RowDgv.Cells(Me.ColRT.Index).Value = .RT
                        RowDgv.Cells(Me.ColRW.Index).Value = .RW
                        RowDgv.Cells(Me.ColKelurahan.Index).Value = .Kelurahan
                        RowDgv.Cells(Me.ColKecamatan.Index).Value = .Kecamatan
                        RowDgv.Cells(Me.ColKabupaten.Index).Value = .Kabupaten
                        RowDgv.Cells(Me.ColPropinsi.Index).Value = .Propinsi
                        RowDgv.Cells(Me.ColKodePos.Index).Value = .Kode_Pos
                        RowDgv.Cells(Me.ColTelepon.Index).Value = .Telepon
                        RowDgv.Cells(Me.ColAlamatPajak1.Index).Value = .Alamat_Pajak_1
                        RowDgv.Cells(Me.ColAlamatPajak2.Index).Value = .Alamat_Pajak_2
                        RowDgv.Cells(Me.ColKodeJenisTransaksi.Index).Value = .kode_jenis_transaksi_pajak

                        'update NPWP 16 Digit.
                        If String.IsNullOrEmpty(.New_NPWP) = False Then
                            RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .New_NPWP
                        Else
                            If CDate(RowDgv.Cells(Me.ColDate.Index).Value).Date >= objNPWP.Tglberlaku Then
                                RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                            Else
                                RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP_Old 'sumber dari master customer 
                            End If
                        End If

                        '//TR : Tambahan baru Nik/Pas
                        'update NPWP 16 Digit.
                        If .Pilih_NIKPASS = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByNIK Then
                            RowDgv.Cells(Me.ColNik.Index).Value = Me.ObjCustomer.GetNoNikPass2(.Primary_Customer, .Pilih_NIKPASS)
                        ElseIf .Pilih_NIKPASS = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByPASSPORT Then
                            RowDgv.Cells(Me.ColPassport.Index).Value = Me.ObjCustomer.GetNoNikPass2(.Primary_Customer, .Pilih_NIKPASS)
                        Else '2025 Core Tax Tahap 2.1
                            RowDgv.Cells(Me.ColOtherID.Index).Value = Me.ObjCustomer.GetNoNikPass2(.Primary_Customer, .Pilih_NIKPASS)
                        End If

                        ''//TR : Tambahan baru Nik/Pas
                        'If .Pilih_NIKPASS = OptionNIKPASS.EnumOptionNikPass.ByNIK Then
                        '    'RowDgv.Cells(Me.ColNik.Index).Value = Me.ObjCustomer.GetNoNikPass2(.Primary_Customer, .Pilih_NIKPASS)
                        '    RowDgv.Cells(Me.ColNik.Index).Value = Me.ObjInv.GetNoNikPassbyInvoice(.Primary)
                        'Else
                        '    'RowDgv.Cells(Me.ColPassport.Index).Value = Me.ObjCustomer.GetNoNikPass2(.Primary_Customer, .Pilih_NIKPASS)
                        '    RowDgv.Cells(Me.ColPassport.Index).Value = Me.ObjInv.GetNoNikPassbyInvoice(.Primary)
                        'End If

                        '2025 Core Tax Tahap 2.1
                        RowDgv.Cells(Me.colNPWPType.Index).Value = .NPWP_Type
                        If RowDgv.Cells(Me.ColNPWPNumber.Index).Value <> "00.000.000.0-000.000" And RowDgv.Cells(Me.ColNPWPNumber.Index).Value <> "-" _
                        And RowDgv.Cells(Me.ColNPWPNumber.Index).Value <> "" And RowDgv.Cells(Me.ColNPWPNumber.Index).Value <> "0000 0000 0000 0000" _
                        And RowDgv.Cells(Me.ColNPWPNumber.Index).Value <> "0000000000000000" And .NPWP_Type <> Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP.Individual Then
                            AdaNPWP = True
                        End If
                        If RowDgv.Cells(Me.ColNik.Index).Value <> "-" And RowDgv.Cells(Me.ColNik.Index).Value <> "" Then
                            AdaNIK = True
                        End If
                        If RowDgv.Cells(Me.ColPassport.Index).Value <> "-" And RowDgv.Cells(Me.ColPassport.Index).Value <> "" Then
                            AdaPASS = True
                        End If
                        '2025 Core Tax Tahap 2.1
                        If RowDgv.Cells(Me.ColOtherID.Index).Value <> "" Then
                            AdaOtherID = True
                        End If

                        If .Tipe_Invoice = TipeSalesInv Then
                            'dp
                            RowDgv.Cells(Me.ColIsDP.Index).Value = 1
                        Else
                            'invoice biasa/pelunasan
                            RowDgv.Cells(Me.ColIsDP.Index).Value = 0
                        End If

                        RowDgv.Cells(Me.ColSPPBNo.Index).Value = .SPPB_No
                        RowDgv.Cells(Me.colEmail.Index).Value = .Email '2025-Coretax
                        RowDgv.Cells(Me.ColIDTKU.Index).Value = .TKU_ID '2025-Coretax

                        '2025 Core Tax Tahap 2.2
                        RowDgv.Cells(Me.colAdditionalInfo.Index).Value = .Additional_Info
                        RowDgv.Cells(Me.ColFacilityStamp.Index).Value = .Cap_Fasilitas
                        RowDgv.Cells(Me.colSPPBDate.Index).Value = .SPPB_Date
                        RowDgv.Cells(Me.colTaxRevised.Index).Value = .Tax_Revised
                        RowDgv.Cells(Me.colTaxType.Index).Value = .Tax_Type
                    End With
                Next
                '//TR : Tambahan baru Nik/Pas
                If AdaNPWP = False Then
                    Me.ColNPWPNumber.Visible = False
                Else
                    Me.ColNPWPNumber.Visible = True
                End If

                If AdaNIK = False Then
                    Me.ColNik.Visible = False
                Else
                    Me.ColNik.Visible = True
                End If
                If AdaPASS = False Then
                    Me.ColPassport.Visible = False
                Else
                    Me.ColPassport.Visible = True
                End If
                '2025 Core Tax Tahap 2.1
                If AdaOtherID = False Then
                    Me.ColOtherID.Visible = False
                Else
                    Me.ColOtherID.Visible = True
                End If
                Me.ColSPPBNo.Visible = True

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

    Private Function getFakturString(ByVal vTaxType As Integer) As String
        Dim retVal As String = ""
        Try
            If vTaxType = 0 Then
                retVal = "01"
            ElseIf vTaxType = 1 Then
                retVal = "02"
            ElseIf vTaxType = 2 Then
                retVal = "03"
            ElseIf vTaxType = 3 Then
                retVal = "04"
            ElseIf vTaxType = 4 Then
                retVal = "05"
            ElseIf vTaxType = 5 Then
                retVal = "06"
            ElseIf vTaxType = 6 Then
                retVal = "07"
            ElseIf vTaxType = 7 Then
                retVal = "08"
            ElseIf vTaxType = 8 Then
                retVal = "09"
            ElseIf vTaxType = 9 Then
                retVal = "10"
            End If
            Return retVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinCrv()
        Dim rpt As New CrpLaporanVAT_2015
        '~ab~
        'Remarks : rapihkan format export excel only data utk laporan vat
        '06/02/2020
        Dim rpt_TipeTransaksi_1 As New CrpLaporanVAT_2015_TipeTransaksi1
        Dim rpt_TipeTransaksi_2 As New CrpLaporanVAT_2015_TipeTransaksi2
        '~

        Dim TBLLap As DseLaporanVAT2025.Tbl_VATDataTable = DS.Tbl_VAT
        Dim RowLap As DseLaporanVAT2025.Tbl_VATRow
        Dim RowDGV As DataGridViewRow
        Dim InOut As Integer = 0
        Dim Tipe As EnumVATType

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
                    RowLap.Tax_Address = .Cells(Me.ColTaxAddress.Index).Value
                    RowLap.Tanggal_Transaksi_Referensi = .Cells(Me.ColTanggalTransaksiRef.Index).Value
                    RowLap.No_Nota_Retur_Pajak = .Cells(Me.ColNoNotaRetur.Index).Value
                    RowLap.Tax_Invoice_No_Referensi = .Cells(Me.ColNoFakturPajakRef.Index).Value
                    RowLap.Nama_Jalan = .Cells(Me.ColNamaJalan.Index).Value
                    RowLap.Blok = .Cells(Me.ColBlok.Index).Value
                    RowLap.Nomor = .Cells(Me.ColNomor.Index).Value
                    RowLap.RT = .Cells(Me.ColRT.Index).Value
                    RowLap.RW = .Cells(Me.ColRW.Index).Value
                    RowLap.Kelurahan = .Cells(Me.ColKelurahan.Index).Value
                    RowLap.Kecamatan = .Cells(Me.ColKecamatan.Index).Value
                    RowLap.Kabupaten = .Cells(Me.ColKabupaten.Index).Value
                    RowLap.Propinsi = .Cells(Me.ColPropinsi.Index).Value
                    RowLap.Kode_Pos = .Cells(Me.ColKodePos.Index).Value
                    RowLap.No_Telepon = .Cells(Me.ColTelepon.Index).Value
                    RowLap.Alamat_Pajak_1 = .Cells(Me.ColAlamatPajak1.Index).Value
                    RowLap.Alamat_Pajak_2 = .Cells(Me.ColAlamatPajak2.Index).Value
                    RowLap.kode_jenis_transaksi_pajak = .Cells(Me.ColKodeJenisTransaksi.Index).Value
                    RowLap.primary_sales_invoice = .Cells(Me.ColPrimaryDocAcuan.Index).Value
                    RowLap.is_dp = .Cells(Me.ColIsDP.Index).Value
                    '//TR : Tambahan baru Nik/Pas
                    RowLap.isNIK = 2
                    If .Cells(Me.ColNik.Index).Value <> "-" And .Cells(Me.ColNik.Index).Value <> "" Then
                        RowLap.NikPass = .Cells(Me.ColNik.Index).Value
                        RowLap.isNIK = 1
                    End If
                    If .Cells(Me.ColPassport.Index).Value <> "-" And .Cells(Me.ColPassport.Index).Value <> "" Then
                        RowLap.NikPass = .Cells(Me.ColPassport.Index).Value
                        RowLap.isNIK = 0
                    End If

                    '2025 Core Tax Tahap 2.1
                    If .Cells(Me.ColOtherID.Index).Value <> "" Then
                        RowLap.NikPass = .Cells(Me.ColOtherID.Index).Value
                        RowLap.isNIK = 2
                    End If
                    RowLap.NPWPType = .Cells(Me.colNPWPType.Index).Value

                    RowLap.SPPB_NO = .Cells(Me.ColSPPBNo.Index).Value
                    RowLap.Email = .Cells(Me.colEmail.Index).Value
                    RowLap.ID_TKU = .Cells(Me.ColIDTKU.Index).Value

                    '2025 Core Tax Tahap 2.2
                    RowLap.Additional_Info = .Cells(Me.colAdditionalInfo.Index).Value
                    RowLap.Facility_Stamp = .Cells(Me.ColFacilityStamp.Index).Value
                    RowLap.Tax_Revised = .Cells(Me.colTaxRevised.Index).Value
                    RowLap.Tax_Type = .Cells(Me.colTaxType.Index).Value
                    RowLap.SPPB_Date = .Cells(Me.colSPPBDate.Index).Value

                    TBLLap.AddTbl_VATRow(RowLap)
                End With
            Next

            'rpt.SetDataSource(DS)
            'rpt.SetParameterValue(rpt.Parameter_Tanggal_Dari.ParameterFieldName, Format(Me.dtpFrom.Value.Date, "dd MMM yyyy"))
            'rpt.SetParameterValue(rpt.Parameter_Tanggal_Sampai.ParameterFieldName, Format(Me.dtpTo.Value.Date, "dd MMM yyyy"))
            'rpt.SetParameterValue(rpt.Parameter_Decimal.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)

            '~ab~
            'Remarks : rapihkan format export excel only data utk laporan vat
            '06/02/2020
            rpt_TipeTransaksi_1.SetDataSource(DS)
            rpt_TipeTransaksi_1.SetParameterValue(rpt_TipeTransaksi_1.Parameter_Tanggal_Dari.ParameterFieldName, Format(Me.dtpFrom.Value.Date, "dd MMM yyyy"))
            rpt_TipeTransaksi_1.SetParameterValue(rpt_TipeTransaksi_1.Parameter_Tanggal_Sampai.ParameterFieldName, Format(Me.dtpTo.Value.Date, "dd MMM yyyy"))
            rpt_TipeTransaksi_1.SetParameterValue(rpt_TipeTransaksi_1.Parameter_Decimal.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)

            rpt_TipeTransaksi_2.SetDataSource(DS)
            rpt_TipeTransaksi_2.SetParameterValue(rpt_TipeTransaksi_2.Parameter_Tanggal_Dari.ParameterFieldName, Format(Me.dtpFrom.Value.Date, "dd MMM yyyy"))
            rpt_TipeTransaksi_2.SetParameterValue(rpt_TipeTransaksi_2.Parameter_Tanggal_Sampai.ParameterFieldName, Format(Me.dtpTo.Value.Date, "dd MMM yyyy"))
            rpt_TipeTransaksi_2.SetParameterValue(rpt_TipeTransaksi_2.Parameter_Decimal.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)
            '~

            Tipe = Me.ComboBox1.SelectedIndex
            'Type Enum hanya 2, Out Sales dan Out Return
            If Tipe = EnumVATType.VAT_OUT_INVOICEDP Then
                'rpt.SetParameterValue(rpt.Parameter_Judul.ParameterFieldName, "VAT Report : VAT Out Purchase Return")
                'rpt.SetParameterValue(rpt.Parameter_Lawan_Transaksi.ParameterFieldName, "Supplier")
                'rpt.SetParameterValue(rpt.Parameter_Tipe_Transaksi.ParameterFieldName, "")
                'rpt.SetParameterValue(rpt.Parameter_TipeTransaksi2.ParameterFieldName, "Purchase Return")
                rpt_TipeTransaksi_2.SetParameterValue(rpt_TipeTransaksi_2.Parameter_Judul.ParameterFieldName, "VAT Report : VAT Out Sales DP")
                rpt_TipeTransaksi_2.SetParameterValue(rpt_TipeTransaksi_2.Parameter_Lawan_Transaksi.ParameterFieldName, "Customer")
                rpt_TipeTransaksi_2.SetParameterValue(rpt_TipeTransaksi_2.Parameter_Tipe_Transaksi.ParameterFieldName, "")
                rpt_TipeTransaksi_2.SetParameterValue(rpt_TipeTransaksi_2.Parameter_TipeTransaksi2.ParameterFieldName, "Sales DP")
                crv1.ReportSource = rpt_TipeTransaksi_2

            ElseIf Tipe = EnumVATType.VAT_OUT_INVOICE Then
                'rpt.SetParameterValue(rpt.Parameter_Judul.ParameterFieldName, "VAT Report : VAT Out Sales")
                'rpt.SetParameterValue(rpt.Parameter_Lawan_Transaksi.ParameterFieldName, "Customer")
                'rpt.SetParameterValue(rpt.Parameter_Tipe_Transaksi.ParameterFieldName, "")
                'rpt.SetParameterValue(rpt.Parameter_TipeTransaksi2.ParameterFieldName, "Sales")
                rpt_TipeTransaksi_2.SetParameterValue(rpt_TipeTransaksi_2.Parameter_Judul.ParameterFieldName, "VAT Report : VAT Out Sales")
                rpt_TipeTransaksi_2.SetParameterValue(rpt_TipeTransaksi_2.Parameter_Lawan_Transaksi.ParameterFieldName, "Customer")
                rpt_TipeTransaksi_2.SetParameterValue(rpt_TipeTransaksi_2.Parameter_Tipe_Transaksi.ParameterFieldName, "")
                rpt_TipeTransaksi_2.SetParameterValue(rpt_TipeTransaksi_2.Parameter_TipeTransaksi2.ParameterFieldName, "Sales")
                crv1.ReportSource = rpt_TipeTransaksi_2
            Else
                'rpt.SetParameterValue(rpt.Parameter_Judul.ParameterFieldName, "VAT Report")
                'rpt.SetParameterValue(rpt.Parameter_Lawan_Transaksi.ParameterFieldName, "Supplier/Customer")
                'rpt.SetParameterValue(rpt.Parameter_Tipe_Transaksi.ParameterFieldName, "Purchase/" + vbCrLf + " Sales")
                'rpt.SetParameterValue(rpt.Parameter_TipeTransaksi2.ParameterFieldName, "")
                rpt_TipeTransaksi_1.SetParameterValue(rpt_TipeTransaksi_1.Parameter_Judul.ParameterFieldName, "VAT Report")
                rpt_TipeTransaksi_1.SetParameterValue(rpt_TipeTransaksi_1.Parameter_Lawan_Transaksi.ParameterFieldName, "Supplier/Customer")
                rpt_TipeTransaksi_1.SetParameterValue(rpt_TipeTransaksi_1.Parameter_Tipe_Transaksi.ParameterFieldName, "Purchase/" + vbCrLf + " Sales")
                rpt_TipeTransaksi_1.SetParameterValue(rpt_TipeTransaksi_1.Parameter_TipeTransaksi2.ParameterFieldName, "")
                crv1.ReportSource = rpt_TipeTransaksi_1
            End If

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
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = -1 Then
                Exit Sub
            End If
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
                        If Me.ObjPR.Apakah_QuickPurchaseReturn(PrimaryDoc) = False Then
                            Me.Show_PR(PrimaryDoc)
                        Else
                            Me.Show_PR_Quick(PrimaryDoc)
                        End If
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
                        If Me.ObjSR.ApakahQuickSR(PrimaryDoc) = False Then
                            Me.Show_SR(PrimaryDoc)
                        Else
                            Me.Show_SR_Quick(PrimaryDoc)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Show_PR_Quick(ByVal vPrimary As Double)
        Try
            'cek apakah user memiliki hak akses yang cukup
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiPurchaseReturn_Quick") = False Then
                MsgBox("Insufficient acces", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Dim ij As Purchase_Module.FrmTransaksiPurchaseReturn_Quick
            Dim j As Windows.Forms.Form

            For Each j In Me.ParentForm.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPurchaseReturn_Quick)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    ij.ShowTransaction(vPrimary)
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiPurchaseReturn_Quick
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

    Private Sub Show_SR_Quick(ByVal vPrimary As Double)
        Try
            'cek apakah user memiliki hak akses yang cukup
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSalesReturn") = False Then
                MsgBox("Insufficient acces", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Dim ij As Sales_Module.FrmTransaksiSalesReturn_Quick
            Dim j As Windows.Forms.Form

            For Each j In Me.ParentForm.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiSalesReturn_Quick)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    ij.ShowTransaction(vPrimary)
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiSalesReturn_Quick
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
        Dim DTA As New DseLaporanVAT2025TableAdapters.DataTable3TableAdapter
        Dim TBL As New DseLaporanVAT2025.DataTable3DataTable
        Dim row As DseLaporanVAT2025.DataTable3Row
        Dim DTA_Supplier As New DseLaporanVAT2025TableAdapters.tblmaster_supplierTableAdapter
        Dim TBL_Supplier As New DseLaporanVAT2025.tblmaster_supplierDataTable
        Dim row_Supplier As DseLaporanVAT2025.tblmaster_supplierRow

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
        Dim DTA_AOP As New DseLaporanVAT2025TableAdapters.DataTable3TableAdapter

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

    Private Sub ComboBox1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.Leave
        '~ab~
        'Remarks : tambah hak akses bisa edit tax number
        '03/02/2020
        Try
            CboIndexBuffer = Me.ComboBox1.SelectedIndex
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            '~ab~
            'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
            '14/01/2020

            If CboIndexBuffer <> Me.ComboBox1.SelectedIndex Then
                Me.dgvItem.Rows.Clear()
                Me.CboIndexBuffer = Me.ComboBox1.SelectedIndex
            End If

            If Me.ComboBox1.SelectedIndex = 0 Or _
            Me.ComboBox1.SelectedIndex = 1 Or _
            Me.ComboBox1.SelectedIndex = 2 Then
                Me.btnCreateCsv2015.Visible = False
                Me.btnCreateCsv2015.TabStop = False
            Else
                Me.btnCreateCsv2015.Visible = True
                Me.btnCreateCsv2015.TabStop = True
            End If

            Me.HideUnhideByCombobox()
            '~

            If Me.ComboBox1.SelectedIndex = EnumVATType.VAT_OUT_INVOICE Or _
            Me.ComboBox1.SelectedIndex = EnumVATType.VAT_OUT_INVOICEDP Then
                Me.BtnCreateCoreTax.Visible = True
                Me.btnImportTaxNumber.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCreateCsv2015_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateCsv2015.Click
        Dim Tipe As FrmLaporanVAT.EnumInOut = FrmLaporanVAT.EnumInOut.Both

        Try
            '2025-Coretax
            Dim expirationDate As Date = New Date(2024, 12, 31)
            If Me.dtpFrom.Enabled = True Then
                If Me.dtpFrom.Value.Date < expirationDate Then
                    MsgBox("This feature is only available after 31 December 2024.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo.Enabled = True Then
                If Me.dtpTo.Value.Date < expirationDate Then
                    MsgBox("This feature is only available after 31 December 2024.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            'end 2025-Coretax

            Tipe = Me.ComboBox1.SelectedIndex
            If Tipe = FrmLaporanVAT.EnumInOut.VATIn_Purchase Then
                Me.ExportPajakMasukan()
            ElseIf Tipe = FrmLaporanVAT.EnumInOut.VATIn_SalesReturn Then
                Me.ExportPajakMasukan_SalesReturn()
            ElseIf Tipe = FrmLaporanVAT.EnumInOut.VATOut_PurchaseReturn Then
                Me.ExportPajakKeluaran_PurchaseReturn()
            ElseIf Tipe = FrmLaporanVAT.EnumInOut.VATOut_Sales Then
                Me.ExportPajakKeluaran()
            End If

            If Me.ChkShowDiscount.Checked = True Then
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.VAT, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ShowDiskon, "1")
            Else
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.VAT, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ShowDiskon, "0")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ExportPajakMasukan_SalesReturn()
        Dim ValTxt As String = "RM;NPWP;NAMA;KD_JENIS_TRANSAKSI;FG_PENGGANTI;NOMOR_FAKTUR;TANGGAL_FAKTUR;IS_CREDITABLE;NOMOR_DOKUMEN_RETUR;TANGGAL_RETUR;MASA_PAJAK_RETUR;TAHUN_PAJAK_RETUR;NILAI_RETUR_DPP;NILAI_RETUR_PPN;NILAI_RETUR_PPNBM"
        Dim RowLaporan As DseLaporanVAT2025.Tbl_VATRow
        Dim DummyIn As String = ""
        Dim DummyOut As String = ""
        Dim Dummy As String = ""

        Dim NoFakturDummy As String = ""
        Dim NoNPWP As String = ""

        Try
            'keluarkan datanya kembali
            'Me.btnShow.PerformClick()
            '~ab~
            'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
            '03/02/2020 
            Dim tblDummy As New DseLaporanVAT2025.Tbl_VATDataTable
            Dim rowDummy As DseLaporanVAT2025.Tbl_VATRow

            Dim ListExport As New List(Of Double)
            Dim ListExportStatus As New List(Of Double)
            ListExport.Clear()
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If row.Cells(Me.ColExport.Index).Value = Boolean.TrueString Then 'jika betul2 belum pernah di export
                    If ListExport.Contains(row.Cells(Me.ColPrimaryDocAcuan.Index).Value) = False Then
                        ListExport.Add(row.Cells(Me.ColPrimaryDocAcuan.Index).Value)
                    End If
                Else
                    If String.IsNullOrEmpty(row.Cells(Me.ColExportStatus.Index).Value) = False Then ' jika sudah pernah di export
                        If ListExportStatus.Contains(row.Cells(Me.ColPrimaryDocAcuan.Index).Value) = False Then
                            ListExportStatus.Add(row.Cells(Me.ColPrimaryDocAcuan.Index).Value)
                        End If
                    End If
                End If
            Next
            '~

            'cek berkaitan dgn no NPWP 
            For Each RowLaporan In DS.Tbl_VAT.Rows
                If ListExport.Contains(RowLaporan.primary_sales_invoice) = True Then
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
                End If
            Next

            If DummyIn <> "" Or DummyOut <> "" Then
                Dummy = "Please provide correct NPWP number for the following document:"
                If DummyIn <> "" Then
                    Dummy = Dummy + vbCrLf + DummyIn
                End If
                If DummyOut <> "" Then
                    Dummy = Dummy + vbCrLf + DummyOut
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
                    If ListExport.Contains(RowLaporan.primary_sales_invoice) = True Then
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
                    End If
                End With
            Next

            If DummyIn <> "" Or DummyOut <> "" Then
                Dummy = "Please provide correct tax invoice number for the following document:"
                If DummyIn <> "" Then
                    Dummy = Dummy + vbCrLf + DummyIn
                End If
                If DummyOut <> "" Then
                    Dummy = Dummy + vbCrLf + DummyOut
                End If

                MsgBox(Dummy, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '~ab~
            'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
            '03/02/2020
            If ListExport.Count = 0 Then
                MsgBox("Please check export minimum one item.", MsgBoxStyle.Information, "Information")
                Me.btnShow.PerformClick()
                Exit Sub
            End If

            tblDummy.Rows.Clear()
            tblDummy = DS.Tbl_VAT.Copy
            For i As Integer = tblDummy.Rows.Count - 1 To 0 Step -1
                rowDummy = tblDummy.Rows(i)
                If ListExport.Contains(rowDummy.primary_sales_invoice) = False Then
                    tblDummy.Rows.Remove(rowDummy)
                End If
            Next

            'If DS.Tbl_VAT.Rows.Count > 0 Then
            If tblDummy.Rows.Count > 0 Then
                If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor

                    For Each rowDummy In tblDummy.Rows
                        'For Each RowLaporan In DS.Tbl_VAT.Rows
                        'With RowLaporan
                        With rowDummy
                            NoFakturDummy = .Tax_Invoice_No
                            NoFakturDummy = NoFakturDummy.Trim
                            NoFakturDummy = NoFakturDummy.Replace("-", "")
                            NoFakturDummy = NoFakturDummy.Replace(".", "")
                            NoNPWP = .NPWP
                            NoNPWP = NoNPWP.Trim
                            NoNPWP = NoNPWP.Replace("-", "")
                            NoNPWP = NoNPWP.Replace(".", "")

                            .Cust_Supp_Name = .Cust_Supp_Name.Trim.Replace(";", "")
                            .Invoice_No = .Invoice_No.Trim.Replace(";", "")
                            .NPWP = .NPWP.Trim.Replace(";", "")
                            .Retur = .Retur.Trim.Replace(";", "")
                            .Tax_Address = .Tax_Address.Trim.Replace(";", "")
                            .Tax_Invoice_No = .Tax_Invoice_No.Trim.Replace(";", "")
                            .Type = .Type.Trim.Replace(";", "")

                            .Cust_Supp_Name = .Cust_Supp_Name.Trim.Replace("""", "")
                            .Invoice_No = .Invoice_No.Trim.Replace("""", "")
                            .NPWP = .NPWP.Trim.Replace("""", "")
                            .Retur = .Retur.Trim.Replace("""", "")
                            .Tax_Address = .Tax_Address.Trim.Replace("""", "")
                            .Tax_Invoice_No = .Tax_Invoice_No.Trim.Replace("""", "")
                            .Type = .Type.Trim.Replace("""", "")

                            ValTxt = ValTxt + vbCrLf + "RM" + ";" + NoNPWP + ";" + .Cust_Supp_Name + ";" + NoFakturDummy.Substring(0, 2).ToCharArray + ";" + NoFakturDummy.Substring(2, 1) + ";" + NoFakturDummy.Substring(3) + _
                            ";" + Format(CDate(.Tanggal_Transaksi_Referensi), "dd/MM/yyyy") + ";1;" _
                            + .No_Nota_Retur_Pajak + ";" + Format(CDate(._Date), "dd/MM/yyyy") + ";" + CDate(._Date).Month.ToString + ";" + CDate(._Date).Year.ToString + ";" _
                            + CStr(Math.Floor(CDbl(.Purchase_Sales))) + ";" + CStr(Math.Floor(CDbl(.VAT))) + ";0"

                        End With
                    Next
                    My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, ValTxt, False)
                    MsgBox("Csv file created. Please be aware that characters used in import process (;,"") will be omitted.", MsgBoxStyle.Information, "Information")

                    For Each RowLaporan In DS.Tbl_VAT.Rows
                        '~ab~
                        'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
                        '03/02/2020
                        With RowLaporan
                            If ListExport.Contains(.primary_sales_invoice) = True Then
                                Me.UpdateExportStatus_InSalesReturn(Me.UserCode & " on " & Format(Now, "dd MMM yyyy - HH.mm"), .primary_sales_invoice)
                            Else
                                If ListExportStatus.Contains(.primary_sales_invoice) = False Then
                                    Me.UpdateExportStatus_InSalesReturn("-", .primary_sales_invoice)
                                End If
                            End If
                        End With
                    Next
                End If
            End If

            '~ab~
            'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
            '03/02/2020
            Me.btnShow.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ExportPajakKeluaran_PurchaseReturn()
        Dim ValTxt As String = "RK;NPWP;NAMA;KD_JENIS_TRANSAKSI;FG_PENGGANTI;NOMOR_FAKTUR;TANGGAL_FAKTUR;NOMOR_DOKUMEN_RETUR;TANGGAL_RETUR;MASA_PAJAK_RETUR;TAHUN_PAJAK_RETUR;NILAI_RETUR_DPP;NILAI_RETUR_PPN;NILAI_RETUR_PPNBM"
        Dim RowLaporan As DseLaporanVAT2025.Tbl_VATRow
        Dim DummyIn As String = ""
        Dim DummyOut As String = ""
        Dim Dummy As String = ""

        Dim NoFakturDummy As String = ""
        Dim NoNPWP As String = ""

        Try
            'keluarkan datanya kembali
            'Me.btnShow.PerformClick()

            '~ab~
            'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
            '03/02/2020 
            Dim tblDummy As New DseLaporanVAT2025.Tbl_VATDataTable
            Dim rowDummy As DseLaporanVAT2025.Tbl_VATRow

            Dim ListExport As New List(Of Double)
            Dim ListExportStatus As New List(Of Double)
            ListExport.Clear()
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If row.Cells(Me.ColExport.Index).Value = Boolean.TrueString Then 'jika betul2 belum pernah di export
                    If ListExport.Contains(row.Cells(Me.ColPrimaryDocAcuan.Index).Value) = False Then
                        ListExport.Add(row.Cells(Me.ColPrimaryDocAcuan.Index).Value)
                    End If
                Else
                    If String.IsNullOrEmpty(row.Cells(Me.ColExportStatus.Index).Value) = False Then ' jika sudah pernah di export
                        If ListExportStatus.Contains(row.Cells(Me.ColPrimaryDocAcuan.Index).Value) = False Then
                            ListExportStatus.Add(row.Cells(Me.ColPrimaryDocAcuan.Index).Value)
                        End If
                    End If
                End If
            Next
            '~

            'cek berkaitan dgn no NPWP 
            For Each RowLaporan In DS.Tbl_VAT.Rows
                If ListExport.Contains(RowLaporan.primary_sales_invoice) = True Then
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
                End If
            Next

            If DummyIn <> "" Or DummyOut <> "" Then
                Dummy = "Please provide correct NPWP number for the following document:"
                If DummyIn <> "" Then
                    Dummy = Dummy + vbCrLf + DummyIn
                End If
                If DummyOut <> "" Then
                    Dummy = Dummy + vbCrLf + DummyOut
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
                    If ListExport.Contains(RowLaporan.primary_sales_invoice) = True Then
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
                    End If
                End With
            Next

            If DummyIn <> "" Or DummyOut <> "" Then
                Dummy = "Please provide correct tax invoice number for the following document:"
                If DummyIn <> "" Then
                    Dummy = Dummy + vbCrLf + DummyIn
                End If
                If DummyOut <> "" Then
                    Dummy = Dummy + vbCrLf + DummyOut
                End If

                MsgBox(Dummy, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '~ab~
            'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
            '03/02/2020
            If ListExport.Count = 0 Then
                MsgBox("Please check export minimum one item.", MsgBoxStyle.Information, "Information")
                Me.btnShow.PerformClick()
                Exit Sub
            End If

            tblDummy.Rows.Clear()
            tblDummy = DS.Tbl_VAT.Copy
            For i As Integer = tblDummy.Rows.Count - 1 To 0 Step -1
                rowDummy = tblDummy.Rows(i)
                If ListExport.Contains(rowDummy.primary_sales_invoice) = False Then
                    tblDummy.Rows.Remove(rowDummy)
                End If
            Next


            'If DS.Tbl_VAT.Rows.Count > 0 Then
            If tblDummy.Rows.Count > 0 Then
                If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor

                    'For Each RowLaporan In DS.Tbl_VAT.Rows
                    For Each rowDummy In tblDummy.Rows
                        With rowDummy
                            NoFakturDummy = .Tax_Invoice_No_Referensi
                            NoFakturDummy = NoFakturDummy.Trim
                            NoFakturDummy = NoFakturDummy.Replace("-", "")
                            NoFakturDummy = NoFakturDummy.Replace(".", "")
                            NoNPWP = .NPWP
                            NoNPWP = NoNPWP.Trim
                            NoNPWP = NoNPWP.Replace("-", "")
                            NoNPWP = NoNPWP.Replace(".", "")

                            .Cust_Supp_Name = .Cust_Supp_Name.Trim.Replace(";", "")
                            .Invoice_No = .Invoice_No.Trim.Replace(";", "")
                            .NPWP = .NPWP.Trim.Replace(";", "")
                            .Retur = .Retur.Trim.Replace(";", "")
                            .Tax_Address = .Tax_Address.Trim.Replace(";", "")
                            .Tax_Invoice_No = .Tax_Invoice_No.Trim.Replace(";", "")
                            .Type = .Type.Trim.Replace(";", "")

                            .Cust_Supp_Name = .Cust_Supp_Name.Trim.Replace("""", "")
                            .Invoice_No = .Invoice_No.Trim.Replace("""", "")
                            .NPWP = .NPWP.Trim.Replace("""", "")
                            .Retur = .Retur.Trim.Replace("""", "")
                            .Tax_Address = .Tax_Address.Trim.Replace("""", "")
                            .Tax_Invoice_No = .Tax_Invoice_No.Trim.Replace("""", "")
                            .Type = .Type.Trim.Replace("""", "")

                            ValTxt = ValTxt + vbCrLf + "RK" + ";" + NoNPWP + ";" + .Cust_Supp_Name + ";" + NoFakturDummy.Substring(0, 2).ToCharArray + ";" + NoFakturDummy.Substring(2, 1) + ";" + NoFakturDummy.Substring(3) + _
                            ";" + Format(CDate(.Tanggal_Transaksi_Referensi), "dd/MM/yyyy") + ";" _
                            + .No_Nota_Retur_Pajak + ";" + Format(CDate(._Date), "dd/MM/yyyy") + ";" + CDate(._Date).Month.ToString + ";" + CDate(._Date).Year.ToString + ";" _
                            + CStr(Math.Floor(CDbl(.Purchase_Sales))) + ";" + CStr(Math.Floor(CDbl(.VAT))) + ";0"

                        End With
                    Next
                    My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, ValTxt, False)
                    MsgBox("Csv file created. Please be aware that characters used in import process (;,"") will be omitted.", MsgBoxStyle.Information, "Information")

                    For Each RowLaporan In DS.Tbl_VAT.Rows
                        '~ab~
                        'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
                        '03/02/2020
                        With RowLaporan
                            If ListExport.Contains(.primary_sales_invoice) = True Then
                                Me.UpdateExportStatus_PurchaseReturn(Me.UserCode & " on " & Format(Now, "dd MMM yyyy - HH.mm"), .primary_sales_invoice)
                            Else
                                If ListExportStatus.Contains(.primary_sales_invoice) = False Then
                                    Me.UpdateExportStatus_PurchaseReturn("-", .primary_sales_invoice)
                                End If
                            End If
                        End With
                    Next
                End If
            End If

            Me.btnShow.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KurangkanItemReturn()
        Dim Filter As String = ""
        Dim row1 As DseLaporanVAT2025.Item_InvoiceRow
        Dim i As Integer = 0

        Try
            For Each rowReturn In Me.TBLReturn.Rows
                Filter = "[Primary Invoice] =" + rowReturn.Primary_Sales_Invoice.ToString + " AND [Primary Product] =" + _
                rowReturn.Primary_Product.ToString + " AND [Primary SO] =" + rowReturn.Primary_SO.ToString + _
                " AND [Primary DO] =" + rowReturn.Primary_DO.ToString
                Me.TBL.DefaultView.RowFilter = Filter

                i = 0
                While Me.TBL.DefaultView.Count > i And rowReturn.Total_Qty > 0
                    row1 = Me.TBL.DefaultView(i).Row
                    If row1.Total_Qty >= rowReturn.Total_Qty Then
                        row1.Total_Qty -= rowReturn.Total_Qty
                        rowReturn.Total_Qty = 0
                    Else
                        rowReturn.Total_Qty -= row1.Total_Qty
                        row1.Total_Qty = 0
                    End If

                    If row1.Total_Qty = 0 Then
                        Me.TBL.Rows.Remove(row1)
                    End If
                End While
            Next

            Me.TBL.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '~ab~
    'Remarks : tambah hak akses bisa edit tax number
    '03/02/2020
    Dim CboIndexBuffer As Integer = -1
    Dim Buffer As Boolean = False
    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        ''~ab~
        ''Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
        ''14/01/2020
        Dim TypeInOut As Integer = 0
        Try
            Me.BisaEditTaxNo = Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmLaporanVAT2_Edit_FP_AOP")
            If e.ColumnIndex = Me.ColExport.Index Or e.ColumnIndex = Me.ColTaxInvoiceNumber.Index Then
                If e.ColumnIndex = Me.ColTaxInvoiceNumber.Index Then
                    TypeInOut = Me.ComboBox1.SelectedIndex
                    If TypeInOut = FrmLaporanVAT.EnumInOut.VATIn_Purchase Then
                        If Me.BisaEditTaxNo = True Then
                            Me.BufferTax = Me.dgvItem.Item(Me.ColTaxInvoiceNumber.Index, e.RowIndex).Value
                        Else
                            e.Cancel = True
                            Exit Sub
                        End If
                    Else
                        e.Cancel = True
                    End If
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HideUnhideByCombobox()
        '~ab~
        'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
        '03/02/2020
        Dim TypeInOut As Integer = 1
        Try
            TypeInOut = Me.ComboBox1.SelectedIndex
            If TypeInOut = EnumVATType.VAT_OUT_INVOICE Or TypeInOut = EnumVATType.VAT_OUT_INVOICEDP Then
                Me.ColExport.Visible = True
                Me.ColExportStatus.Visible = True
            Else
                Me.ColExport.Visible = False
                Me.ColExportStatus.Visible = False
            End If



        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateExportStatus_Invoice(ByVal vExportStatus As String, ByVal vPrimaryInvoice As Double)
        '~ab~
        'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
        '03/02/2020
        Dim DTA As New DseLaporanVAT2025TableAdapters.InvoiceTableAdapter
        Try
            If vPrimaryInvoice > -1 Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTA.UpdateQueryExportStatus(vExportStatus, vPrimaryInvoice)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateExportStatus_PurchaseReturn(ByVal vExportStatus As String, ByVal vPrimaryPurchase As Double)
        '~ab~
        'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
        '03/02/2020
        Dim DTA As New DseLaporanVAT2025TableAdapters.Purchase_ReturnTableAdapter
        Try
            If vPrimaryPurchase > -1 Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTA.UpdateQueryExportStatus(vExportStatus, vPrimaryPurchase)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateExportStatus_InPurchase(ByVal vExportStatus As String, ByVal vPrimaryPurchase As Double)
        '~ab~
        'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
        '03/02/2020
        Dim DTA As New DseLaporanVAT2025TableAdapters.TBL_AOPTableAdapter
        Try
            If vPrimaryPurchase > -1 Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTA.UpdateQueryExportStatus(vExportStatus, vPrimaryPurchase)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateExportStatus_InSalesReturn(ByVal vExportStatus As String, ByVal vPrimarySR As Double)
        '~ab~
        'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
        '03/02/2020
        Dim DTA As New DseLaporanVAT2025TableAdapters.Sales_ReturnTableAdapter
        Try
            If vPrimarySR > -1 Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTA.UpdateQueryExportStatus(vExportStatus, vPrimarySR)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahNoFakturBenar(ByVal TxtTaxInvoiceNumber As String, ByVal ExportStatus As String) As Boolean
        '~ab~
        'Remarks : penambahan batasan khusus utk no inv 011 dan 071 
        '03/02/2020 

        Try
            TxtTaxInvoiceNumber = TxtTaxInvoiceNumber.Trim
            If String.IsNullOrEmpty(ExportStatus) = False And TxtTaxInvoiceNumber.StartsWith("011") = True Then
                Return False
            End If
            If String.IsNullOrEmpty(ExportStatus) = False And TxtTaxInvoiceNumber.StartsWith("071") = True Then
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub LlbCheck_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles Llbcheck.LinkClicked
        '~ab~
        'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
        '03/02/2020
        Try
            Me.Cursor = Cursors.WaitCursor
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If Me.ApakahNoFakturBenar(row.Cells(Me.ColTaxInvoiceNumber.Index).Value, row.Cells(Me.ColExportStatus.Index).Value) = True Then
                    row.Cells(Me.ColExport.Index).Value = Boolean.TrueString
                End If
            Next
            Me.Cursor = Cursors.Arrow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        '~ab~
        'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
        '03/02/2020
        Try
            Me.Cursor = Cursors.WaitCursor
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                row.Cells(Me.ColExport.Index).Value = Boolean.FalseString
            Next
            Me.Cursor = Cursors.Arrow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim ListOfNoInvoice As New List(Of String)
    Private Function ApakahDataBenar() As Boolean
        Dim TxtTaxInvoiceNumber As String
        Dim OK As Boolean = True
        Try
            '~ab~
            'Remarks : penambahan batasan khusus utk no inv 011 dan 071 
            '03/02/2020 

            ListOfNoInvoice.Clear()
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If row.Cells(Me.ColExport.Index).Value = Boolean.TrueString Then
                    If Me.ApakahNoFakturBenar(row.Cells(Me.ColTaxInvoiceNumber.Index).Value, row.Cells(Me.ColExportStatus.Index).Value) = False Then
                        If Me.ListOfNoInvoice.Contains(row.Cells(Me.ColInvoiceNumber.Index).Value) = False Then
                            Me.ListOfNoInvoice.Add(row.Cells(Me.ColInvoiceNumber.Index).Value)
                        End If
                    End If
                End If
            Next

            If Me.ListOfNoInvoice.Count > 0 Then
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            '~ab~
            'Remarks : penambahan batasan khusus utk no inv 011 dan 071 
            '03/02/2020 

            If e.ColumnIndex = Me.ColExport.Index Then
                If Me.ComboBox1.SelectedIndex = FrmLaporanVAT.EnumInOut.VATOut_Sales And Me.dgvItem.Item(Me.ColExport.Index, e.RowIndex).Value = Boolean.TrueString Then
                    If Me.ApakahNoFakturBenar(Me.dgvItem.Item(Me.ColTaxInvoiceNumber.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColExportStatus.Index, e.RowIndex).Value) = False Then
                        Dim Hasil As MsgBoxResult
                        Hasil = MsgBox("Tax invoice No:" & Me.dgvItem.Item(Me.ColTaxInvoiceNumber.Index, e.RowIndex).Value & " has been exported, will it continue?", MsgBoxStyle.YesNo, "Confirmation")
                        If Hasil = MsgBoxResult.No Then
                            Me.dgvItem.Item(Me.ColExport.Index, e.RowIndex).Value = Boolean.FalseString
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCreateCoreTax.Click
        Dim RowLaporan As DseLaporanVAT2025.Tbl_VATRow
        Dim DummyIn As String = ""
        Dim DummyOut As String = ""
        Dim Dummy As String = ""
        Dim KodeJenisTransaksiPajakDummy As String = ""
        Dim NoFakturDummy As String = ""
        Dim NoNPWP As String = ""
        Dim DPP_DP As Decimal = 0
        Dim PPN_DP As Decimal = 0
        Dim DefPrecIDR As Integer = 0
        Dim FG_UANG_MUKA As Integer = 0
        Dim DPRef As String = ""
        Dim DummyDPP As Decimal = 0
        Dim DummyPPN As Decimal = 0
        Dim ItemCount As Integer = 0
        Dim Dummy1 As Decimal = 0


        '#by ahdar | 19/06/20 | @terkait dpp yang di export dari laporan vat mengambil sub total + pengiriman
        Dim OngkosKirim As Decimal = 0
        Dim IsShipmentWithVAT As Boolean = False

        Dim hj As New FrmTransaksiFakturPajak
        Try
            '2025-Coretax
            Dim expirationDate As Date = New Date(2024, 12, 31)
            If Me.dtpFrom.Enabled = True Then
                If Me.dtpFrom.Value.Date <= expirationDate Then
                    MsgBox("This feature is only available after 1 January 2025.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo.Enabled = True Then
                If Me.dtpTo.Value.Date <= expirationDate Then
                    MsgBox("This feature is only available after 1 January 2025.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            'end 2025-Coretax

            '~ab~
            'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
            '03/02/2020 
            Dim tblDummy As New DseLaporanVAT2025.Tbl_VATDataTable
            Dim rowDummy As DseLaporanVAT2025.Tbl_VATRow
            Dim rowDummy1 As DseLaporanVAT2025.Tbl_VATRow

            Dim ListExport As New List(Of Double)
            Dim ListExportStatus As New List(Of Double)
            ListExport.Clear()
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                If row.Cells(Me.ColExport.Index).Value = Boolean.TrueString Then 'jika betul2 belum pernah di export
                    If ListExport.Contains(row.Cells(Me.ColPrimaryDocAcuan.Index).Value) = False Then
                        ListExport.Add(row.Cells(Me.ColPrimaryDocAcuan.Index).Value)
                    End If
                Else
                    If String.IsNullOrEmpty(row.Cells(Me.ColExportStatus.Index).Value) = False Then ' jika sudah pernah di export
                        If ListExportStatus.Contains(row.Cells(Me.ColPrimaryDocAcuan.Index).Value) = False Then
                            ListExportStatus.Add(row.Cells(Me.ColPrimaryDocAcuan.Index).Value)
                        End If
                    End If
                End If
            Next
            '~

            DefPrecIDR = Me.ObjCur.GetDecimal_DefaultCurrency
            'keluarkan datanya kembali
            'Me.btnShow.PerformClick()

            'cek berkaitan dgn no NPWP 
            For Each RowLaporan In DS.Tbl_VAT.Rows
                If ListExport.Contains(RowLaporan.primary_sales_invoice) = True Then
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
                End If
            Next

            If DummyIn <> "" Or DummyOut <> "" Then
                Dummy = "Please provide correct NPWP number for the following document:"
                If DummyIn <> "" Then
                    Dummy = Dummy + vbCrLf + DummyIn
                End If
                If DummyOut <> "" Then
                    Dummy = Dummy + vbCrLf + DummyOut
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
                    If ListExport.Contains(RowLaporan.primary_sales_invoice) = True Then
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
                    End If
                End With
            Next
            If DummyIn <> "" Or DummyOut <> "" Then
                Dummy = "Please provide correct tax invoice number for the following document:"
                If DummyIn <> "" Then
                    Dummy = Dummy + vbCrLf + DummyIn
                End If
                If DummyOut <> "" Then
                    Dummy = Dummy + vbCrLf + DummyOut
                End If

                MsgBox(Dummy, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '~ab~
            'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
            '03/02/2020
            If ListExport.Count = 0 Then
                MsgBox("Please check export minimum one item.", MsgBoxStyle.Information, "Information")
                Me.btnShow.PerformClick()
                Exit Sub
            End If

            tblDummy.Rows.Clear()
            'tblDummy = DS.Tbl_VAT.Copy
            'For i As Integer = tblDummy.Rows.Count - 1 To 0 Step -1
            '    rowDummy = tblDummy.Rows(i)
            '    If ListExport.Contains(rowDummy.primary_sales_invoice) = False Then
            '        tblDummy.Rows.Remove(rowDummy)
            '    End If
            'Next

            For Each PrimaryInvoice As Double In ListExport
                DS.Tbl_VAT.DefaultView.RowFilter = "[primary sales invoice] =" & PrimaryInvoice
                If DS.Tbl_VAT.DefaultView.Count > 0 Then
                    rowDummy = DS.Tbl_VAT.DefaultView(0).Row
                    rowDummy1 = tblDummy.NewTbl_VATRow
                    rowDummy1.ItemArray = rowDummy.ItemArray
                    tblDummy.Rows.Add(rowDummy1)
                End If
            Next
            '~

            'If DS.Tbl_VAT.Rows.Count > 0 Then
            If tblDummy.Rows.Count > 0 Then
                hj.tblDummy = tblDummy
                hj.showDiscount = Me.ChkShowDiscount.Checked
                hj.dtpFrom = Me.dtpFrom.Value.Date
                hj.dtpTo = Me.dtpTo.Value.Date
                hj.BackColorAll = Me.BackColorAll
                hj.ShowDialog()

                If hj.Berhasil = True Then
                    For Each RowLaporan In DS.Tbl_VAT.Rows
                        '~ab~
                        'Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
                        '03/02/2020
                        With RowLaporan
                            If ListExport.Contains(.primary_sales_invoice) = True Then
                                Me.UpdateExportStatus_Invoice(Me.UserCode & " on " & Format(Now, "dd MMM yyyy - HH.mm"), .primary_sales_invoice)
                            Else
                                If ListExportStatus.Contains(.primary_sales_invoice) = False Then
                                    Me.UpdateExportStatus_Invoice("-", .primary_sales_invoice)
                                End If
                            End If
                        End With
                    Next
                End If
            End If

            Me.btnShow.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImportTaxNumber.Click
        Dim dataValid As Double = 0
        Dim hj As New FrmLaporanVAT2025_Import
        Dim tblDummy As New DseLaporanVAT2025.Tbl_VATDataTable
        Try
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("No data can be imported", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            tblDummy.Rows.Clear()
            tblDummy = DS.Tbl_VAT.Copy

            hj.BackColorAll = Me.BackColorAll
            hj.UserCode = Me.UserCode
            hj.tblDummy = tblDummy
            'hj.TblPurchaseReturn = Me.TBLOut_PR
            Control.CheckForIllegalCrossThreadCalls = False
            hj.ShowDialog()
            Me.btnShow.PerformClick()

            'If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then

            '    If Me.LoadCSV(Me.OpenFileDialog1.FileName) Then
            '        Me.TxtFilePath.Text = Me.OpenFileDialog1.FileName
            '        dataValid = Me.FilterTaxNo()
            '        If dataValid > 0 Then
            '            Me.CheckReference(0, Me.TblImport.Rows.Count - 1, True)
            '            MsgBox(dataValid & " Data Found", MsgBoxStyle.Information, "Information")
            '        Else
            '            MsgBox("Data not found.", MsgBoxStyle.Information, "Information")
            '        End If
            '    Else
            '        Me.TxtFilePath.Text = ""
            '    End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class