''inf-001
Option Explicit On
Option Strict Off

Imports Microsoft.VisualBasic.FileIO

Public Class FrmLaporanVAT2
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
    Dim ObjPR As New ClsTransaksiPurchaseReturn

    Dim DTAin As New DseLaporanVAT1TableAdapters.DataTable1TableAdapter
    Dim TBLin As New DseLaporanVAT1.DataTable1DataTable
    Dim DTAIN_Purchase As New DseLaporanVAT1TableAdapters.TBL_AOPTableAdapter
    Dim TBLIN_Purchase As New DseLaporanVAT1.TBL_AOPDataTable
    Dim rowIN_Purchase As DseLaporanVAT1.TBL_AOPRow
    Dim DTAIN_SR As New DseLaporanVAT1TableAdapters.Sales_ReturnTableAdapter
    Dim TBLIN_SR As New DseLaporanVAT1.Sales_ReturnDataTable
    Dim rowIN_SR As DseLaporanVAT1.Sales_ReturnRow

    Dim DTAOut As New DseLaporanVAT1TableAdapters.DataTable2TableAdapter
    Dim TBLOut As New DseLaporanVAT1.DataTable2DataTable
    Dim DTAOut_Sales As New DseLaporanVAT1TableAdapters.InvoiceTableAdapter
    Dim TBLOut_Sales As New DseLaporanVAT1.InvoiceDataTable
    Dim rowOut_Sales As DseLaporanVAT1.InvoiceRow
    Dim DTAReturn As New DseLaporanVAT1TableAdapters.Item_Invoice_ReturnTableAdapter
    Dim TBLReturn As New DseLaporanVAT1.Item_Invoice_ReturnDataTable
    Dim rowReturn As DseLaporanVAT1.Item_Invoice_ReturnRow
    Dim DTA As New DseLaporanVAT1TableAdapters.Item_InvoiceTableAdapter
    Dim TBL As New DseLaporanVAT1.Item_InvoiceDataTable
    Dim row As DseLaporanVAT1.Item_InvoiceRow
    Dim DTAItemDP As New DseLaporanVAT1TableAdapters.Item_Invoice_DPTableAdapter
    Dim TBLItemDP As New DseLaporanVAT1.Item_Invoice_DPDataTable
    Dim rowItemDP As DseLaporanVAT1.Item_Invoice_DPRow

    Dim DTAOut_PR As New DseLaporanVAT1TableAdapters.Purchase_ReturnTableAdapter
    Dim TBLOut_PR As New DseLaporanVAT1.Purchase_ReturnDataTable
    Dim rowOut_PR As DseLaporanVAT1.Purchase_ReturnRow
    Dim DTAOut_PR_Ref As New DseLaporanVAT1TableAdapters.Purchase_Return_ReferenceTableAdapter
    Dim TBLOut_PR_Ref As New DseLaporanVAT1.Purchase_Return_ReferenceDataTable
    Dim rowOut_PR_Ref As DseLaporanVAT1.Purchase_Return_ReferenceRow
    Dim ObjInv As New Sales_Module.ClsTransaksiSalesInvoice
    Dim ObjAOP As New ClsTransaksiAOP
    Dim ObjSR As New Sales_Module.ClsTransaksiSalesReturn
    Dim OptionNIKPASS As Core_Module.clsMasterCustomer
    Dim ObjCustomer As New Core_Module.clsMasterCustomer
    Dim ObjLicense As New Core_Module.ClsLicense
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting

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
        VAT_IN_AOP
        VAT_IN_AOP_DP
        VAT_IN_QUICK_AOP
        VAT_IN_RETURN
        VAT_OUT_INVOICE
        VAT_OUT_DP
        VAT_OUT_QUICK_INVOICE
        VAT_OUT_RETURN
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
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
        Dim RowLaporan As DseLaporanVAT1.Tbl_VATRow
        Dim DummyIn As String = ""
        Dim DummyOut As String = ""
        Dim Dummy As String = ""

        Dim NoFakturDummy As String = ""
        Dim NoNPWP As String = ""

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
                    Dummy = Dummy + vbCrLf + DummyIn
                End If
                If DummyOut <> "" Then
                    Dummy = Dummy + vbCrLf + DummyOut
                End If

                MsgBox(Dummy, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If DS.Tbl_VAT.Rows.Count > 0 Then
                If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor

                    For Each RowLaporan In DS.Tbl_VAT.Rows
                        With RowLaporan
                            NoFakturDummy = RowLaporan.Tax_Invoice_No
                            NoFakturDummy = NoFakturDummy.Trim
                            NoFakturDummy = NoFakturDummy.Replace("-", "")
                            NoFakturDummy = NoFakturDummy.Replace(".", "")
                            NoNPWP = RowLaporan.NPWP
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
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ExportPajakKeluaran()
        Dim ValTxt As String = "FK;KD_JENIS_TRANSAKSI;FG_PENGGANTI;NOMOR_FAKTUR;MASA_PAJAK;TAHUN_PAJAK;TANGGAL_FAKTUR;NPWP;NAMA;ALAMAT_LENGKAP;JUMLAH_DPP;JUMLAH_PPN;JUMLAH_PPNBM;ID_KETERANGAN_TAMBAHAN;FG_UANG_MUKA;UANG_MUKA_DPP;UANG_MUKA_PPN;UANG_MUKA_PPNBM;REFERENSI (NIK/PASSPORT)"  '//TR : Tambahan baru Nik/Pas
        ValTxt = ValTxt + vbCrLf + "LT;NPWP;NAMA;JALAN;BLOK;NOMOR;RT;RW;KECAMATAN;KELURAHAN;KABUPATEN;PROPINSI;KODE_POS;NO_TELEPON"
        ValTxt = ValTxt + vbCrLf + "OF;KODE_OBJEK;NAMA;HARGA_SATUAN;JUMLAH_BARANG;HARGA_TOTAL;DISKON;DPP;PPN;TARIF_PPNBM;PPNBM"

        Dim RowLaporan As DseLaporanVAT1.Tbl_VATRow
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

        Try
            DefPrecIDR = Me.ObjCur.GetDecimal_DefaultCurrency
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
                    Dummy = Dummy + vbCrLf + DummyIn
                End If
                If DummyOut <> "" Then
                    Dummy = Dummy + vbCrLf + DummyOut
                End If

                MsgBox(Dummy, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If DS.Tbl_VAT.Rows.Count > 0 Then
                If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor

                    'load data item
                    Me.DTA.Fill(Me.TBL, Me.dtpTo.Value, Me.dtpFrom.Value)
                    Me.DTAReturn.Fill(Me.TBLReturn, Me.dtpTo.Value, Me.dtpFrom.Value)
                    Me.KurangkanItemReturn()
                    Me.DTAItemDP.Fill(Me.TBLItemDP, Me.dtpTo.Value, Me.dtpFrom.Value)

                    For Each RowLaporan In DS.Tbl_VAT.Rows
                        With RowLaporan
                            KodeJenisTransaksiPajakDummy = .kode_jenis_transaksi_pajak
                            If KodeJenisTransaksiPajakDummy = "-1" Then
                                KodeJenisTransaksiPajakDummy = ""
                            Else
                                KodeJenisTransaksiPajakDummy = CInt(.kode_jenis_transaksi_pajak) + 1
                            End If
                            NoFakturDummy = RowLaporan.Tax_Invoice_No
                            NoFakturDummy = NoFakturDummy.Trim
                            NoFakturDummy = NoFakturDummy.Replace("-", "")
                            NoFakturDummy = NoFakturDummy.Replace(".", "")
                            NoNPWP = RowLaporan.NPWP
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
                            Me.TBLItemDP.DefaultView.RowFilter = "[Primary Sales Invoice] = " + RowLaporan.primary_sales_invoice.ToString
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
                                .Invoice_No & " " & .NikPass
                                'lawan transaksi
                                ValTxt = ValTxt + vbCrLf + "LT" + ";" + NoNPWP + ";" + .Cust_Supp_Name + ";" + .Nama_Jalan + ";" + .Blok + ";" + .Nomor + ";" + _
                                .RT + ";" + .RW + ";" + .Kecamatan + ";" + .Kelurahan + ";" + .Kabupaten + ";" + .Propinsi + ";" + .Kode_Pos + ";" + .No_Telepon
                                'item

                                ItemCount = 0
                                For Each dv As DataRowView In Me.TBL.DefaultView
                                    ItemCount += 1
                                    Me.row = dv.Row
                                    If ItemCount = Me.TBL.DefaultView.Count Then
                                        Me.row.DPP_Value = DummyDPP
                                        Me.row.PPN_Value = DummyPPN
                                    Else
                                        Me.row.DPP_Value = Me.row.Sales_Amount
                                        Me.row.PPN_Value = Math.Round(Me.row.Sales_Amount * (Me.row.Tax / 100), DefPrecIDR, MidpointRounding.AwayFromZero)
                                        DummyDPP -= Me.row.DPP_Value
                                        DummyPPN -= Me.row.PPN_Value
                                    End If
                                Next
                                For Each dv As DataRowView In Me.TBL.DefaultView
                                    Me.row = dv.Row
                                    If Me.row.price_by_multi_unit = 0 Then
                                        Me.row.Total_Qty_Multi = Me.row.Total_Qty
                                    End If
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
                                .Invoice_No & " " & .NikPass  '//TR : Tambahan baru Nik/Pas
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
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub ExportPajakKeluaran()
    '    Dim ValTxt As String = "FK;KD_JENIS_TRANSAKSI;FG_PENGGANTI;NOMOR_FAKTUR;MASA_PAJAK;TAHUN_PAJAK;TANGGAL_FAKTUR;NPWP;NAMA;ALAMAT_LENGKAP;JUMLAH_DPP;JUMLAH_PPN;JUMLAH_PPNBM;ID_KETERANGAN_TAMBAHAN;FG_UANG_MUKA;UANG_MUKA_DPP;UANG_MUKA_PPN;UANG_MUKA_PPNBM;REFERENSI"
    '    ValTxt = ValTxt + vbCrLf + "LT;NPWP;NAMA;JALAN;BLOK;NOMOR;RT;RW;KECAMATAN;KELURAHAN;KABUPATEN;PROPINSI;KODE_POS;NO_TELEPON"
    '    ValTxt = ValTxt + vbCrLf + "OF;KODE_OBJEK;NAMA;HARGA_SATUAN;JUMLAH_BARANG;HARGA_TOTAL;DISKON;DPP;PPN;TARIF_PPNBM;PPNBM"

    '    Dim RowLaporan As DseLaporanVAT1.Tbl_VATRow
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

        Dim RowIn As DseLaporanVAT1.DataTable1Row
        Dim RowOut As DseLaporanVAT1.DataTable2Row

        Dim RowDgv As DataGridViewRow
        Dim PrimaryGRNMin As Double = Double.MaxValue
        Dim PrimaryGRNMax As Double = Double.MinValue
        Dim TipeSalesInv As Sales_Module.ClsTransaksiSalesInvoice.EnumTipeInvoice = Sales_Module.ClsTransaksiSalesInvoice.EnumTipeInvoice.InvoiceDP
        Me.ColNoNotaRetur.Visible = False
        Dim AdaNPWP As Boolean = False
        Dim AdaNIK As Boolean = False
        Dim AdaPASS As Boolean = False
        Dim EndLicenseDate As Date = #1/1/1950#
        'Update NPWP 16 Digit
        Dim objNPWP As New Core_Module.ClsTrans_SettingNPWPNo
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
                If Me.dtpFrom.Value.Date > expirationDate Then
                    MsgBox("This feature is no longer available after 31 December 2024.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo.Enabled = True Then
                If Me.dtpTo.Value.Date > expirationDate Then
                    MsgBox("This feature is no longer available after 31 December 2024.", _
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

            If TypeInOut = FrmLaporanVAT.EnumInOut.VATIn Then
                Me.DTAin.Fill(TBLin, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipePR)
                For Each RowIn In TBLin.Rows
                    With RowIn
                        'If .VAT_dipungut = 0 Then
                        '    SalerOrPr = 0
                        '    VAT = 0
                        'Else : .VAT_dipungut = 1
                        SalerOrPr = .Purchase
                        VAT = .VAT
                        'End If
                        Me.dgvItem.Rows.Add()
                        RowDgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        RowDgv.Cells(Me.ColTipeTransaksiString.Index).Value = ._In
                        RowDgv.Cells(Me.ColSuppCust.Index).Value = .Supp
                        'Update NPWP 16 Digit
                        'RowDgv.Cells(Me.ColDate.Index).Value = Format(.Supplier_Invoice_Date.Date, "dd MMM yyyy")
                        If .Tipe_AOP = 0 And .Retur = 0 Then
                            RowDgv.Cells(Me.ColDate.Index).Value = Format(.Supplier_Invoice_Date.Date, "dd MMM yyyy")
                        Else
                            RowDgv.Cells(Me.ColDate.Index).Value = Format(.Tanggal_sampai.Date, "dd MMM yyyy")
                        End If

                        'Update NPWP 16 Digit
                        'RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                        If String.IsNullOrEmpty(.New_NPWP) = False Then
                            RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .New_NPWP
                        Else
                            If CDate(RowDgv.Cells(Me.ColDate.Index).Value).Date >= objNPWP.Tglberlaku Then
                                RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                            Else
                                RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP_Old
                            End If
                        End If
                        If .Retur = 0 Then
                            RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no2
                        Else
                            RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no
                        End If
                        RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .Tax_invoice_no
                        RowDgv.Cells(Me.ColPurchaseSales.Index).Value = Me.ObjCD.DigitGrouping(SalerOrPr.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(VAT.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColPrimaryDocAcuan.Index).Value = .Primary
                        RowDgv.Cells(Me.ColEnumType.Index).Value = .Retur
                        RowDgv.Cells(Me.ColEnumInOut.Index).Value = FrmLaporanVAT.EnumInOut.VATIn
                        RowDgv.Cells(Me.ColTaxAddress.Index).Value = ""
                        RowDgv.Cells(Me.ColTanggalTransaksiRef.Index).Value = Format(#1/1/1950#, "dd MMM yyyy")
                        RowDgv.Cells(Me.ColNoNotaRetur.Index).Value = ""
                        RowDgv.Cells(Me.ColNoFakturPajakRef.Index).Value = ""
                        RowDgv.Cells(Me.ColNamaJalan.Index).Value = ""
                        RowDgv.Cells(Me.ColBlok.Index).Value = ""
                        RowDgv.Cells(Me.ColNomor.Index).Value = ""
                        RowDgv.Cells(Me.ColRT.Index).Value = ""
                        RowDgv.Cells(Me.ColRW.Index).Value = ""
                        RowDgv.Cells(Me.ColKelurahan.Index).Value = ""
                        RowDgv.Cells(Me.ColKecamatan.Index).Value = ""
                        RowDgv.Cells(Me.ColKabupaten.Index).Value = ""
                        RowDgv.Cells(Me.ColPropinsi.Index).Value = ""
                        RowDgv.Cells(Me.ColKodePos.Index).Value = ""
                        RowDgv.Cells(Me.ColTelepon.Index).Value = ""
                        RowDgv.Cells(Me.ColAlamatPajak1.Index).Value = ""
                        RowDgv.Cells(Me.ColAlamatPajak2.Index).Value = ""
                        RowDgv.Cells(Me.ColKodeJenisTransaksi.Index).Value = ""
                        RowDgv.Cells(Me.ColIsDP.Index).Value = 0
                    End With
                Next
                Me.SalinCrv()
            ElseIf TypeInOut = FrmLaporanVAT.EnumInOut.VATOut Then
                Me.DTAOut.Fill(TBLOut, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipeSR)
                For Each RowOut In Me.TBLOut
                    With RowOut
                        'If .VAT_Dipungut = 0 Then
                        '    SalerOrPr = 0
                        '    VAT = 0
                        'Else : .VAT_Dipungut = 1
                        SalerOrPr = .Sales
                        VAT = .VAT
                        'End If
                        Me.dgvItem.Rows.Add()
                        RowDgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        RowDgv.Cells(Me.ColTipeTransaksiString.Index).Value = .Out
                        RowDgv.Cells(Me.ColSuppCust.Index).Value = .Cust
                        'Update NPWP 16 Digit
                        'RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                        If String.IsNullOrEmpty(.New_NPWP) = False Then
                            RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .New_NPWP
                        Else
                            If CDate(RowDgv.Cells(Me.ColDate.Index).Value).Date >= objNPWP.Tglberlaku Then
                                RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                            Else
                                RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP_Old
                            End If
                        End If
                        RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no
                        RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .No_Faktur_Pajak
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
                        RowDgv.Cells(Me.ColNamaJalan.Index).Value = ""
                        RowDgv.Cells(Me.ColBlok.Index).Value = ""
                        RowDgv.Cells(Me.ColNomor.Index).Value = ""
                        RowDgv.Cells(Me.ColRT.Index).Value = ""
                        RowDgv.Cells(Me.ColRW.Index).Value = ""
                        RowDgv.Cells(Me.ColKelurahan.Index).Value = ""
                        RowDgv.Cells(Me.ColKecamatan.Index).Value = ""
                        RowDgv.Cells(Me.ColKabupaten.Index).Value = ""
                        RowDgv.Cells(Me.ColPropinsi.Index).Value = ""
                        RowDgv.Cells(Me.ColKodePos.Index).Value = ""
                        RowDgv.Cells(Me.ColTelepon.Index).Value = ""
                        RowDgv.Cells(Me.ColAlamatPajak1.Index).Value = ""
                        RowDgv.Cells(Me.ColAlamatPajak2.Index).Value = ""
                        RowDgv.Cells(Me.ColKodeJenisTransaksi.Index).Value = ""
                        RowDgv.Cells(Me.ColIsDP.Index).Value = 0

                        'Update NPWP 16 Digit
                        If .Pilih_NIKPASS = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByNIK Then
                            RowDgv.Cells(Me.ColNik.Index).Value = .No_NIKPASS
                        Else
                            RowDgv.Cells(Me.ColPassport.Index).Value = .No_NIKPASS
                        End If
                    End With
                Next
                Me.SalinCrv()
            ElseIf TypeInOut = FrmLaporanVAT.EnumInOut.Both Then
                Me.DTAin.Fill(TBLin, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipePR)
                For Each RowIn In TBLin.Rows
                    With RowIn
                        'If .VAT_dipungut = 0 Then
                        '    SalerOrPr = 0
                        '    VAT = 0
                        'Else : .VAT_dipungut = 1
                        SalerOrPr = .Purchase
                        VAT = .VAT
                        'End If
                        Me.dgvItem.Rows.Add()
                        RowDgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        RowDgv.Cells(Me.ColTipeTransaksiString.Index).Value = ._In
                        RowDgv.Cells(Me.ColSuppCust.Index).Value = .Supp
                        'Update NPWP 16 Digit
                        'RowDgv.Cells(Me.ColDate.Index).Value = Format(.Supplier_Invoice_Date.Date, "dd MMM yyyy")
                        If .Tipe_AOP = 0 And .Retur = 0 Then
                            RowDgv.Cells(Me.ColDate.Index).Value = Format(.Supplier_Invoice_Date.Date, "dd MMM yyyy")
                        Else
                            RowDgv.Cells(Me.ColDate.Index).Value = Format(.Tanggal_sampai.Date, "dd MMM yyyy")
                        End If

                        'Update NPWP 16 Digit
                        'RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                        If String.IsNullOrEmpty(.New_NPWP) = False Then
                            RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .New_NPWP
                        Else
                            If CDate(RowDgv.Cells(Me.ColDate.Index).Value).Date >= objNPWP.Tglberlaku Then
                                RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                            Else
                                RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP_Old
                            End If
                        End If
                        If .Retur = 0 Then
                            RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no2
                        Else
                            RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no
                        End If
                        RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .Tax_invoice_no
                        RowDgv.Cells(Me.ColPurchaseSales.Index).Value = Me.ObjCD.DigitGrouping(SalerOrPr.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(VAT.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColPrimaryDocAcuan.Index).Value = .Primary
                        RowDgv.Cells(Me.ColEnumType.Index).Value = .Retur
                        RowDgv.Cells(Me.ColEnumInOut.Index).Value = FrmLaporanVAT.EnumInOut.VATIn
                        RowDgv.Cells(Me.ColTaxAddress.Index).Value = ""
                        RowDgv.Cells(Me.ColTanggalTransaksiRef.Index).Value = Format(#1/1/1950#, "dd MMM yyyy")
                        RowDgv.Cells(Me.ColNoNotaRetur.Index).Value = ""
                        RowDgv.Cells(Me.ColNoFakturPajakRef.Index).Value = ""
                        RowDgv.Cells(Me.ColNamaJalan.Index).Value = ""
                        RowDgv.Cells(Me.ColBlok.Index).Value = ""
                        RowDgv.Cells(Me.ColNomor.Index).Value = ""
                        RowDgv.Cells(Me.ColRT.Index).Value = ""
                        RowDgv.Cells(Me.ColRW.Index).Value = ""
                        RowDgv.Cells(Me.ColKelurahan.Index).Value = ""
                        RowDgv.Cells(Me.ColKecamatan.Index).Value = ""
                        RowDgv.Cells(Me.ColKabupaten.Index).Value = ""
                        RowDgv.Cells(Me.ColPropinsi.Index).Value = ""
                        RowDgv.Cells(Me.ColKodePos.Index).Value = ""
                        RowDgv.Cells(Me.ColTelepon.Index).Value = ""
                        RowDgv.Cells(Me.ColAlamatPajak1.Index).Value = ""
                        RowDgv.Cells(Me.ColAlamatPajak2.Index).Value = ""
                        RowDgv.Cells(Me.ColKodeJenisTransaksi.Index).Value = ""
                        RowDgv.Cells(Me.ColIsDP.Index).Value = 0
                    End With
                Next
                Me.DTAOut.Fill(TBLOut, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipeSR)
                For Each RowOut In Me.TBLOut
                    With RowOut
                        'If .VAT_Dipungut = 0 Then
                        '    SalerOrPr = 0
                        '    VAT = 0
                        'Else : .VAT_Dipungut = 1
                        SalerOrPr = .Sales
                        VAT = .VAT
                        'End If
                        Me.dgvItem.Rows.Add()
                        RowDgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        RowDgv.Cells(Me.ColTipeTransaksiString.Index).Value = .Out
                        RowDgv.Cells(Me.ColSuppCust.Index).Value = .Cust
                        'Update NPWP 16 Digit
                        'RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                        If String.IsNullOrEmpty(.New_NPWP) = False Then
                            RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .New_NPWP
                        Else
                            If CDate(RowDgv.Cells(Me.ColDate.Index).Value).Date >= objNPWP.Tglberlaku Then
                                RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                            Else
                                RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP_Old
                            End If
                        End If
                        RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no
                        RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .No_Faktur_Pajak
                        RowDgv.Cells(Me.ColDate.Index).Value = Format(.Tanggal_dari.Date, "dd MMM yyyy")
                        RowDgv.Cells(Me.ColPurchaseSales.Index).Value = Me.ObjCD.DigitGrouping(SalerOrPr.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(VAT.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColPrimaryDocAcuan.Index).Value = .Primary
                        RowDgv.Cells(Me.ColEnumType.Index).Value = .Retur
                        RowDgv.Cells(Me.ColEnumInOut.Index).Value = FrmLaporanVAT.EnumInOut.VATOut
                        RowDgv.Cells(Me.ColTanggalTransaksiRef.Index).Value = Format(#1/1/1950#, "dd MMM yyyy")
                        RowDgv.Cells(Me.ColNoNotaRetur.Index).Value = ""
                        RowDgv.Cells(Me.ColTaxAddress.Index).Value = ""
                        RowDgv.Cells(Me.ColNoFakturPajakRef.Index).Value = ""
                        RowDgv.Cells(Me.ColNamaJalan.Index).Value = ""
                        RowDgv.Cells(Me.ColBlok.Index).Value = ""
                        RowDgv.Cells(Me.ColNomor.Index).Value = ""
                        RowDgv.Cells(Me.ColRT.Index).Value = ""
                        RowDgv.Cells(Me.ColRW.Index).Value = ""
                        RowDgv.Cells(Me.ColKelurahan.Index).Value = ""
                        RowDgv.Cells(Me.ColKecamatan.Index).Value = ""
                        RowDgv.Cells(Me.ColKabupaten.Index).Value = ""
                        RowDgv.Cells(Me.ColPropinsi.Index).Value = ""
                        RowDgv.Cells(Me.ColKodePos.Index).Value = ""
                        RowDgv.Cells(Me.ColTelepon.Index).Value = ""
                        RowDgv.Cells(Me.ColAlamatPajak1.Index).Value = ""
                        RowDgv.Cells(Me.ColAlamatPajak2.Index).Value = ""
                        RowDgv.Cells(Me.ColKodeJenisTransaksi.Index).Value = ""
                        RowDgv.Cells(Me.ColIsDP.Index).Value = 0

                        'Update NPWP 16 Digit
                        If .Pilih_NIKPASS = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByNIK Then
                            RowDgv.Cells(Me.ColNik.Index).Value = .No_NIKPASS
                        Else
                            RowDgv.Cells(Me.ColPassport.Index).Value = .No_NIKPASS
                        End If

                    End With
                Next
                Me.SalinCrv()
            ElseIf TypeInOut = FrmLaporanVAT.EnumInOut.VATIn_Purchase Then
                Me.ColPurchaseSales.HeaderText = "Purchase"
                Me.ColSuppCust.HeaderText = "Supplier"
                Me.DTAIN_Purchase.Fill(TBLIN_Purchase, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
                For Each rowIN_Purchase In TBLIN_Purchase.Rows
                    With rowIN_Purchase
                        'If .VAT_dipungut = 0 Then
                        '    SalerOrPr = 0
                        '    VAT = 0
                        'Else : .VAT_dipungut = 1
                        SalerOrPr = .Purchase
                        VAT = .VAT
                        'End If
                        Me.dgvItem.Rows.Add()
                        RowDgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)

                        ''~ab~
                        ''Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
                        ''03/02/2020
                        'If String.IsNullOrEmpty(.Export_Status) = False Then
                        '    RowDgv.Cells(Me.ColExportStatus.Index).Value = .Export_Status
                        'Else
                        '    RowDgv.Cells(Me.ColExportStatus.Index).Value = ""
                        'End If
                        ''~

                        RowDgv.Cells(Me.ColTipeTransaksiString.Index).Value = ._In
                        RowDgv.Cells(Me.ColSuppCust.Index).Value = .Supp
                        RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP_No
                        If .Retur = 0 Then
                            RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no2
                        Else
                            RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no
                        End If
                        RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .Tax_invoice_no
                        'Update NPWP 16 Digit
                        'RowDgv.Cells(Me.ColDate.Index).Value = Format(.Supplier_Invoice_Date.Date, "dd MMM yyyy")
                        If .Tipe_AOP = 0 Then
                            RowDgv.Cells(Me.ColDate.Index).Value = Format(.Supplier_Invoice_Date.Date, "dd MMM yyyy")
                        Else
                            RowDgv.Cells(Me.ColDate.Index).Value = Format(.Tanggal_sampai.Date, "dd MMM yyyy")
                        End If
                        RowDgv.Cells(Me.ColPurchaseSales.Index).Value = Me.ObjCD.DigitGrouping(SalerOrPr.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(VAT.ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColPrimaryDocAcuan.Index).Value = .Primary
                        RowDgv.Cells(Me.ColEnumType.Index).Value = .Retur
                        RowDgv.Cells(Me.ColEnumInOut.Index).Value = FrmLaporanVAT.EnumInOut.VATIn
                        RowDgv.Cells(Me.ColTaxAddress.Index).Value = .tax_addres
                        RowDgv.Cells(Me.ColNoFakturPajakRef.Index).Value = ""
                        RowDgv.Cells(Me.ColNamaJalan.Index).Value = ""
                        RowDgv.Cells(Me.ColBlok.Index).Value = ""
                        RowDgv.Cells(Me.ColNomor.Index).Value = ""
                        RowDgv.Cells(Me.ColRT.Index).Value = ""
                        RowDgv.Cells(Me.ColRW.Index).Value = ""
                        RowDgv.Cells(Me.ColKelurahan.Index).Value = ""
                        RowDgv.Cells(Me.ColKecamatan.Index).Value = ""
                        RowDgv.Cells(Me.ColKabupaten.Index).Value = ""
                        RowDgv.Cells(Me.ColPropinsi.Index).Value = ""
                        RowDgv.Cells(Me.ColKodePos.Index).Value = ""
                        RowDgv.Cells(Me.ColTelepon.Index).Value = ""
                        RowDgv.Cells(Me.ColAlamatPajak1.Index).Value = ""
                        RowDgv.Cells(Me.ColAlamatPajak2.Index).Value = ""
                        RowDgv.Cells(Me.ColKodeJenisTransaksi.Index).Value = ""
                        RowDgv.Cells(Me.ColIsDP.Index).Value = 0
                    End With
                Next
                Me.SalinCrv()
            ElseIf TypeInOut = FrmLaporanVAT.EnumInOut.VATIn_SalesReturn Then
                Me.ColNoNotaRetur.Visible = True
                Me.ColInvoiceNumber.HeaderText = "Sales Return Number"
                Me.ColTaxInvoiceNumber.HeaderText = "Ref. Tax Invoice Number"
                Me.ColPurchaseSales.HeaderText = "Sales Return"
                Me.ColSuppCust.HeaderText = "Customer"

                Me.DTAIN_SR.Fill(TBLIN_SR, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipeSR, 0)
                For Each rowIN_SR In Me.TBLIN_SR
                    With rowIN_SR
                        'If .VAT_Dipungut = 0 Then
                        '    SalerOrPr = 0
                        '    VAT = 0
                        'Else : .VAT_Dipungut = 1
                        SalerOrPr = .Sales
                        VAT = .VAT
                        'End If
                        Me.dgvItem.Rows.Add()
                        RowDgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)

                        ''~ab~
                        ''Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
                        ''24/01/2020
                        'If String.IsNullOrEmpty(.Export_Status) = False Then
                        '    RowDgv.Cells(Me.ColExportStatus.Index).Value = .Export_Status
                        'Else
                        '    RowDgv.Cells(Me.ColExportStatus.Index).Value = ""
                        'End If
                        ''~

                        RowDgv.Cells(Me.ColTipeTransaksiString.Index).Value = .Out
                        RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no
                        RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .No_Faktur_Pajak
                        RowDgv.Cells(Me.ColDate.Index).Value = Format(.Tanggal_dari.Date, "dd MMM yyyy")
                        RowDgv.Cells(Me.ColPurchaseSales.Index).Value = Me.ObjCD.DigitGrouping((SalerOrPr * -1).ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping((VAT * -1).ToString, Me.DefPrecBase)
                        RowDgv.Cells(Me.ColPrimaryDocAcuan.Index).Value = .Primary
                        RowDgv.Cells(Me.ColEnumType.Index).Value = .Retur
                        RowDgv.Cells(Me.ColTaxAddress.Index).Value = .tax_address
                        RowDgv.Cells(Me.ColEnumInOut.Index).Value = FrmLaporanVAT.EnumInOut.VATOut
                        RowDgv.Cells(Me.ColNoNotaRetur.Index).Value = .No_Nota_Retur_Pajak
                        RowDgv.Cells(Me.ColNoFakturPajakRef.Index).Value = ""
                        RowDgv.Cells(Me.ColNamaJalan.Index).Value = ""
                        RowDgv.Cells(Me.ColBlok.Index).Value = ""
                        RowDgv.Cells(Me.ColNomor.Index).Value = ""
                        RowDgv.Cells(Me.ColRT.Index).Value = ""
                        RowDgv.Cells(Me.ColRW.Index).Value = ""
                        RowDgv.Cells(Me.ColKelurahan.Index).Value = ""
                        RowDgv.Cells(Me.ColKecamatan.Index).Value = ""
                        RowDgv.Cells(Me.ColKabupaten.Index).Value = ""
                        RowDgv.Cells(Me.ColPropinsi.Index).Value = ""
                        RowDgv.Cells(Me.ColKodePos.Index).Value = ""
                        RowDgv.Cells(Me.ColTelepon.Index).Value = ""
                        RowDgv.Cells(Me.ColAlamatPajak1.Index).Value = ""
                        RowDgv.Cells(Me.ColAlamatPajak2.Index).Value = ""
                        RowDgv.Cells(Me.ColKodeJenisTransaksi.Index).Value = ""
                        RowDgv.Cells(Me.ColIsDP.Index).Value = 0
                        If .Primary_Customer_Quick <> -1 Then
                            RowDgv.Cells(Me.ColNoFakturPajakRef.Index).Value = .Tax_Invoice_No_Quick
                            RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .Tax_Invoice_No_Quick
                            RowDgv.Cells(Me.ColTanggalTransaksiRef.Index).Value = Format(.Invoice_Date_Quick, "dd MMM yyyy")
                            RowDgv.Cells(Me.ColSuppCust.Index).Value = .Cust_Name_Quick
                            'Update NPWP 16 Digit
                            'RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP_Np_Quick
                            If String.IsNullOrEmpty(.No_NPWP) = False Then
                                RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .No_NPWP 'sumber dari tblsalesreturn
                            Else
                                If CDate(RowDgv.Cells(Me.ColDate.Index).Value).Date >= objNPWP.Tglberlaku Then
                                    RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP_Np_Quick 'sumber dari master customer 
                                Else
                                    RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP_Old_Quick 'sumber dari master customer 
                                End If
                            End If

                            'Update NPWP 16 Digit
                            If .Pilih_NIKPASS_Quick = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByNIK Then
                                RowDgv.Cells(Me.ColNik.Index).Value = .No_NIKPASS_Quick
                            Else
                                RowDgv.Cells(Me.ColPassport.Index).Value = .No_NIKPASS_Quick
                            End If
                        Else
                            RowDgv.Cells(Me.ColSuppCust.Index).Value = .Cust
                            RowDgv.Cells(Me.ColTanggalTransaksiRef.Index).Value = Format(.Tanggal_Transaksi_Referensi, "dd MMM yyyy")
                            'Update NPWP 16 Digit
                            'RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                            If String.IsNullOrEmpty(.No_NPWP) = False Then
                                RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .No_NPWP 'sumber dari tblsalesreturn
                            Else
                                If CDate(RowDgv.Cells(Me.ColDate.Index).Value).Date >= objNPWP.Tglberlaku Then
                                    RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP 'sumber dari master customer
                                Else
                                    RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP_Old 'sumber dari master customer 
                                End If
                            End If

                            'Update NPWP 16 Digit
                            If .Pilih_NIKPASS = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByNIK Then
                                RowDgv.Cells(Me.ColNik.Index).Value = .No_NIKPASS
                            Else
                                RowDgv.Cells(Me.ColPassport.Index).Value = .No_NIKPASS
                            End If
                        End If
                    End With
                Next
                Me.SalinCrv()
            ElseIf TypeInOut = FrmLaporanVAT.EnumInOut.VATOut_Sales Then
                Me.ColPurchaseSales.HeaderText = "Sales"
                Me.ColSuppCust.HeaderText = "Customer"
                Me.DTAOut_Sales.Fill(TBLOut_Sales, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
                For Each rowOut_Sales In Me.TBLOut_Sales
                    With rowOut_Sales
                        'If .VAT_Dipungut = 0 Then
                        '    SalerOrPr = 0
                        '    VAT = 0
                        'Else : .VAT_Dipungut = 1
                        SalerOrPr = .Sales
                        VAT = .VAT1
                        'End If
                        Me.dgvItem.Rows.Add()
                        RowDgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)

                        ''~ab~
                        ''Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
                        ''14/01/2020
                        'If String.IsNullOrEmpty(.Export_Status) = False Then
                        '    RowDgv.Cells(Me.ColExportStatus.Index).Value = .Export_Status
                        'Else
                        '    RowDgv.Cells(Me.ColExportStatus.Index).Value = ""
                        'End If
                        ''~

                        RowDgv.Cells(Me.ColTipeTransaksiString.Index).Value = .Out
                        RowDgv.Cells(Me.ColSuppCust.Index).Value = .Cust
                        RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no
                        RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .No_Faktur_Pajak
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
                        'RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
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
                        If .Pilih_NIKPASS = OptionNIKPASS.EnumOptionNikPass.ByNIK Then
                            RowDgv.Cells(Me.ColNik.Index).Value = Me.ObjCustomer.GetNoNikPass2(.Primary_Customer, .Pilih_NIKPASS)
                            'RowDgv.Cells(Me.ColNik.Index).Value = Me.ObjInv.GetNoNikPassbyInvoice(.Primary)
                        Else
                            RowDgv.Cells(Me.ColPassport.Index).Value = Me.ObjCustomer.GetNoNikPass2(.Primary_Customer, .Pilih_NIKPASS)
                            'RowDgv.Cells(Me.ColPassport.Index).Value = Me.ObjInv.GetNoNikPassbyInvoice(.Primary)
                        End If

                        If .NPWP <> "00.000.000.0-000.000" And .NPWP <> "-" And .NPWP <> "" Then
                            AdaNPWP = True
                        End If
                        If RowDgv.Cells(Me.ColNik.Index).Value <> "-" And RowDgv.Cells(Me.ColNik.Index).Value <> "" Then
                            AdaNIK = True
                        End If
                        If RowDgv.Cells(Me.ColPassport.Index).Value <> "-" And RowDgv.Cells(Me.ColPassport.Index).Value <> "" Then
                            AdaPASS = True
                        End If

                        If .Tipe_Invoice = TipeSalesInv Then
                            'dp
                            RowDgv.Cells(Me.ColIsDP.Index).Value = 1
                        Else
                            'invoice biasa/pelunasan
                            RowDgv.Cells(Me.ColIsDP.Index).Value = 0
                        End If
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
                Me.SalinCrv()
            ElseIf TypeInOut = FrmLaporanVAT.EnumInOut.VATOut_PurchaseReturn Then
                Me.ColNoNotaRetur.Visible = True
                Me.ColInvoiceNumber.HeaderText = "Purchase Return Number"
                Me.ColPurchaseSales.HeaderText = "Purchase Return"
                Me.ColTaxInvoiceNumber.HeaderText = "Ref. Tax Invoice Number"
                Me.ColSuppCust.HeaderText = "Supplier"
                Me.DTAOut_PR.Fill(TBLOut_PR, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TipePR, 0)
                For Each rowOut_PR In TBLOut_PR.Rows
                    With rowOut_PR
                        If .Primary_GRN > PrimaryGRNMax Then
                            PrimaryGRNMax = .Primary_GRN
                        End If
                        If .Primary_GRN < PrimaryGRNMin Then
                            PrimaryGRNMin = -.Primary_GRN
                        End If
                    End With
                Next
                If TBLOut_PR.Rows.Count > 0 Then
                    Me.DTAOut_PR_Ref.FillBy_Primary(TBLOut_PR_Ref, PrimaryGRNMax, PrimaryGRNMin)
                    For Each rowOut_PR In TBLOut_PR.Rows
                        With rowOut_PR
                            'If .VAT_dipungut = 0 Then
                            '    SalerOrPr = 0
                            '    VAT = 0
                            'Else : .VAT_dipungut = 1
                            SalerOrPr = .Purchase
                            VAT = .VAT
                            'End If
                            Me.dgvItem.Rows.Add()
                            RowDgv = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)

                            ''~ab~
                            ''Remarks : penambahan export status & export checklist maka user bisa memilih export ke csv yang dipilih saja
                            ''03/02/2020
                            'If String.IsNullOrEmpty(.Export_Status) = False Then
                            '    RowDgv.Cells(Me.ColExportStatus.Index).Value = .Export_Status
                            'Else
                            '    RowDgv.Cells(Me.ColExportStatus.Index).Value = ""
                            'End If
                            ''~

                            RowDgv.Cells(Me.ColTipeTransaksiString.Index).Value = ._In
                            RowDgv.Cells(Me.ColSuppCust.Index).Value = .Supp
                            If .Retur = 0 Then
                                RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no2
                            Else
                                RowDgv.Cells(Me.ColInvoiceNumber.Index).Value = .Invoice_no
                            End If
                            RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .Tax_invoice_no
                            RowDgv.Cells(Me.ColDate.Index).Value = Format(.Tanggal_dari.Date, "dd MMM yyyy")
                            RowDgv.Cells(Me.ColPurchaseSales.Index).Value = Me.ObjCD.DigitGrouping((SalerOrPr * -1).ToString, Me.DefPrecBase)
                            RowDgv.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping((VAT * -1).ToString, Me.DefPrecBase)
                            RowDgv.Cells(Me.ColPrimaryDocAcuan.Index).Value = .Primary
                            RowDgv.Cells(Me.ColEnumType.Index).Value = .Retur
                            RowDgv.Cells(Me.ColEnumInOut.Index).Value = FrmLaporanVAT.EnumInOut.VATIn
                            RowDgv.Cells(Me.ColTaxAddress.Index).Value = .tax_address
                            Me.TBLOut_PR_Ref.DefaultView.RowFilter = "[Primary GRN] = " + .Primary_GRN.ToString
                            If Me.TBLOut_PR_Ref.DefaultView.Count = 0 Then
                                RowDgv.Cells(Me.ColTanggalTransaksiRef.Index).Value = .Ref_Inv_Date
                                RowDgv.Cells(Me.ColNoNotaRetur.Index).Value = .Tax_invoice_no
                                RowDgv.Cells(Me.ColNoFakturPajakRef.Index).Value = .Ref_Tax_Inv_No
                                RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = .Ref_Tax_Inv_No
                            Else
                                Me.rowOut_PR_Ref = Me.TBLOut_PR_Ref.DefaultView(0).Row
                                RowDgv.Cells(Me.ColTanggalTransaksiRef.Index).Value = Format(Me.rowOut_PR_Ref.Tanggal_Transaksi, "dd MMM yyyy")
                                RowDgv.Cells(Me.ColNoNotaRetur.Index).Value = .Tax_invoice_no
                                RowDgv.Cells(Me.ColNoFakturPajakRef.Index).Value = Me.rowOut_PR_Ref.Tax_Invoice_Number
                                RowDgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = RowDgv.Cells(Me.ColNoFakturPajakRef.Index).Value
                            End If
                            RowDgv.Cells(Me.ColNamaJalan.Index).Value = ""
                            RowDgv.Cells(Me.ColBlok.Index).Value = ""
                            RowDgv.Cells(Me.ColNomor.Index).Value = ""
                            RowDgv.Cells(Me.ColRT.Index).Value = ""
                            RowDgv.Cells(Me.ColRW.Index).Value = ""
                            RowDgv.Cells(Me.ColKelurahan.Index).Value = ""
                            RowDgv.Cells(Me.ColKecamatan.Index).Value = ""
                            RowDgv.Cells(Me.ColKabupaten.Index).Value = ""
                            RowDgv.Cells(Me.ColPropinsi.Index).Value = ""
                            RowDgv.Cells(Me.ColKodePos.Index).Value = ""
                            RowDgv.Cells(Me.ColTelepon.Index).Value = ""
                            RowDgv.Cells(Me.ColAlamatPajak1.Index).Value = ""
                            RowDgv.Cells(Me.ColAlamatPajak2.Index).Value = ""
                            RowDgv.Cells(Me.ColKodeJenisTransaksi.Index).Value = ""
                            RowDgv.Cells(Me.ColIsDP.Index).Value = 0

                            'update NPWP 16 Digit.
                            'RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                            If String.IsNullOrEmpty(.New_NPWP_No) = False Then
                                RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .New_NPWP_No
                            Else
                                If CDate(RowDgv.Cells(Me.ColDate.Index).Value).Date >= objNPWP.Tglberlaku Then
                                    RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP
                                Else
                                    RowDgv.Cells(Me.ColNPWPNumber.Index).Value = .NPWP_Old
                                End If
                            End If
                        End With
                    Next
                    Me.SalinCrv()
                End If
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
        Dim rpt As New CrpLaporanVAT_2015
        Dim TBLLap As DseLaporanVAT1.Tbl_VATDataTable = DS.Tbl_VAT
        Dim RowLap As DseLaporanVAT1.Tbl_VATRow
        Dim RowDGV As DataGridViewRow
        Dim InOut As Integer = 0
        Dim Tipe As FrmLaporanVAT.EnumInOut

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
                    If .Cells(Me.ColNik.Index).Value <> "-" And .Cells(Me.ColNik.Index).Value <> "" Then
                        RowLap.NikPass = "(NIK: " & .Cells(Me.ColNik.Index).Value & ")"
                    End If
                    If .Cells(Me.ColPassport.Index).Value <> "-" And .Cells(Me.ColPassport.Index).Value <> "" Then
                        RowLap.NikPass = "(PASSPORT: " & .Cells(Me.ColPassport.Index).Value & ")"
                    End If
                    TBLLap.AddTbl_VATRow(RowLap)
                End With
            Next

            rpt.SetDataSource(DS)
            rpt.SetParameterValue(rpt.Parameter_Tanggal_Dari.ParameterFieldName, Format(Me.dtpFrom.Value.Date, "dd MMM yyyy"))
            rpt.SetParameterValue(rpt.Parameter_Tanggal_Sampai.ParameterFieldName, Format(Me.dtpTo.Value.Date, "dd MMM yyyy"))
            rpt.SetParameterValue(rpt.Parameter_Decimal.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)
            Tipe = Me.ComboBox1.SelectedIndex
            If Tipe = FrmLaporanVAT.EnumInOut.VATIn Then
                rpt.SetParameterValue(rpt.Parameter_Judul.ParameterFieldName, "VAT Report : VAT In Total")
                rpt.SetParameterValue(rpt.Parameter_Lawan_Transaksi.ParameterFieldName, "Supplier")
                rpt.SetParameterValue(rpt.Parameter_Tipe_Transaksi.ParameterFieldName, "Purchase/" + vbCrLf + "Purchase Return")
                rpt.SetParameterValue(rpt.Parameter_TipeTransaksi2.ParameterFieldName, "")
            ElseIf Tipe = FrmLaporanVAT.EnumInOut.VATOut Then
                rpt.SetParameterValue(rpt.Parameter_Judul.ParameterFieldName, "VAT Report : VAT Out Total")
                rpt.SetParameterValue(rpt.Parameter_Lawan_Transaksi.ParameterFieldName, "Customer")
                rpt.SetParameterValue(rpt.Parameter_Tipe_Transaksi.ParameterFieldName, "Sales/" + vbCrLf + "Sales Return")
                rpt.SetParameterValue(rpt.Parameter_TipeTransaksi2.ParameterFieldName, "")
            ElseIf Tipe = FrmLaporanVAT.EnumInOut.Both Then
                rpt.SetParameterValue(rpt.Parameter_Judul.ParameterFieldName, "VAT Report")
                rpt.SetParameterValue(rpt.Parameter_Lawan_Transaksi.ParameterFieldName, "Supplier/Customer")
                rpt.SetParameterValue(rpt.Parameter_Tipe_Transaksi.ParameterFieldName, "Purchase/" + vbCrLf + " Sales")
                rpt.SetParameterValue(rpt.Parameter_TipeTransaksi2.ParameterFieldName, "")
            ElseIf Tipe = FrmLaporanVAT.EnumInOut.VATIn_Purchase Then
                rpt.SetParameterValue(rpt.Parameter_Judul.ParameterFieldName, "VAT Report : VAT In Purchase")
                rpt.SetParameterValue(rpt.Parameter_Lawan_Transaksi.ParameterFieldName, "Supplier")
                rpt.SetParameterValue(rpt.Parameter_Tipe_Transaksi.ParameterFieldName, "")
                rpt.SetParameterValue(rpt.Parameter_TipeTransaksi2.ParameterFieldName, "Purchase")
            ElseIf Tipe = FrmLaporanVAT.EnumInOut.VATIn_SalesReturn Then
                rpt.SetParameterValue(rpt.Parameter_Judul.ParameterFieldName, "VAT Report : VAT In Sales Return")
                rpt.SetParameterValue(rpt.Parameter_Lawan_Transaksi.ParameterFieldName, "Customer")
                rpt.SetParameterValue(rpt.Parameter_Tipe_Transaksi.ParameterFieldName, "")
                rpt.SetParameterValue(rpt.Parameter_TipeTransaksi2.ParameterFieldName, "Sales Return")
            ElseIf Tipe = FrmLaporanVAT.EnumInOut.VATOut_PurchaseReturn Then
                rpt.SetParameterValue(rpt.Parameter_Judul.ParameterFieldName, "VAT Report : VAT Out Purchase Return")
                rpt.SetParameterValue(rpt.Parameter_Lawan_Transaksi.ParameterFieldName, "Supplier")
                rpt.SetParameterValue(rpt.Parameter_Tipe_Transaksi.ParameterFieldName, "")
                rpt.SetParameterValue(rpt.Parameter_TipeTransaksi2.ParameterFieldName, "Purchase Return")
            ElseIf Tipe = FrmLaporanVAT.EnumInOut.VATOut_Sales Then
                rpt.SetParameterValue(rpt.Parameter_Judul.ParameterFieldName, "VAT Report : VAT Out Sales")
                rpt.SetParameterValue(rpt.Parameter_Lawan_Transaksi.ParameterFieldName, "Customer")
                rpt.SetParameterValue(rpt.Parameter_Tipe_Transaksi.ParameterFieldName, "")
                rpt.SetParameterValue(rpt.Parameter_TipeTransaksi2.ParameterFieldName, "Sales")
            Else
                rpt.SetParameterValue(rpt.Parameter_Judul.ParameterFieldName, "VAT Report")
                rpt.SetParameterValue(rpt.Parameter_Lawan_Transaksi.ParameterFieldName, "Supplier/Customer")
                rpt.SetParameterValue(rpt.Parameter_Tipe_Transaksi.ParameterFieldName, "Purchase/" + vbCrLf + " Sales")
                rpt.SetParameterValue(rpt.Parameter_TipeTransaksi2.ParameterFieldName, "")
            End If
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

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            If Me.ComboBox1.SelectedIndex = 0 Or _
            Me.ComboBox1.SelectedIndex = 1 Or _
            Me.ComboBox1.SelectedIndex = 2 Then
                Me.btnCreateCsv2015.Visible = False
                Me.btnCreateCsv2015.TabStop = False
            Else
                Me.btnCreateCsv2015.Visible = True
                Me.btnCreateCsv2015.TabStop = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCreateCsv2015_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateCsv2015.Click
        Dim Tipe As FrmLaporanVAT.EnumInOut = FrmLaporanVAT.EnumInOut.Both

        Try
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

            '2025-Coretax
            Dim expirationDate As Date = New Date(2024, 12, 31)
            If Me.dtpFrom.Enabled = True Then
                If Me.dtpFrom.Value.Date > expirationDate Then
                    MsgBox("This feature is no longer available after 31 December 2024.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo.Enabled = True Then
                If Me.dtpTo.Value.Date > expirationDate Then
                    MsgBox("This feature is no longer available after 31 December 2024.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            'end 2025-Coretax
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ExportPajakMasukan_SalesReturn()
        Dim ValTxt As String = "RM;NPWP;NAMA;KD_JENIS_TRANSAKSI;FG_PENGGANTI;NOMOR_FAKTUR;TANGGAL_FAKTUR;IS_CREDITABLE;NOMOR_DOKUMEN_RETUR;TANGGAL_RETUR;MASA_PAJAK_RETUR;TAHUN_PAJAK_RETUR;NILAI_RETUR_DPP;NILAI_RETUR_PPN;NILAI_RETUR_PPNBM"
        Dim RowLaporan As DseLaporanVAT1.Tbl_VATRow
        Dim DummyIn As String = ""
        Dim DummyOut As String = ""
        Dim Dummy As String = ""

        Dim NoFakturDummy As String = ""
        Dim NoNPWP As String = ""

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
                    Dummy = Dummy + vbCrLf + DummyIn
                End If
                If DummyOut <> "" Then
                    Dummy = Dummy + vbCrLf + DummyOut
                End If

                MsgBox(Dummy, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If DS.Tbl_VAT.Rows.Count > 0 Then
                If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor

                    For Each RowLaporan In DS.Tbl_VAT.Rows
                        With RowLaporan
                            NoFakturDummy = RowLaporan.Tax_Invoice_No
                            NoFakturDummy = NoFakturDummy.Trim
                            NoFakturDummy = NoFakturDummy.Replace("-", "")
                            NoFakturDummy = NoFakturDummy.Replace(".", "")
                            NoNPWP = RowLaporan.NPWP
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
                            + RowLaporan.No_Nota_Retur_Pajak + ";" + Format(CDate(._Date), "dd/MM/yyyy") + ";" + CDate(._Date).Month.ToString + ";" + CDate(._Date).Year.ToString + ";" _
                            + CStr(Math.Floor(CDbl(.Purchase_Sales))) + ";" + CStr(Math.Floor(CDbl(.VAT))) + ";0"

                        End With
                    Next
                    My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, ValTxt, False)
                    MsgBox("Csv file created. Please be aware that characters used in import process (;,"") will be omitted.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ExportPajakKeluaran_PurchaseReturn()
        Dim ValTxt As String = "RK;NPWP;NAMA;KD_JENIS_TRANSAKSI;FG_PENGGANTI;NOMOR_FAKTUR;TANGGAL_FAKTUR;NOMOR_DOKUMEN_RETUR;TANGGAL_RETUR;MASA_PAJAK_RETUR;TAHUN_PAJAK_RETUR;NILAI_RETUR_DPP;NILAI_RETUR_PPN;NILAI_RETUR_PPNBM"
        Dim RowLaporan As DseLaporanVAT1.Tbl_VATRow
        Dim DummyIn As String = ""
        Dim DummyOut As String = ""
        Dim Dummy As String = ""

        Dim NoFakturDummy As String = ""
        Dim NoNPWP As String = ""

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
                    Dummy = Dummy + vbCrLf + DummyIn
                End If
                If DummyOut <> "" Then
                    Dummy = Dummy + vbCrLf + DummyOut
                End If

                MsgBox(Dummy, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If DS.Tbl_VAT.Rows.Count > 0 Then
                If Me.SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor

                    For Each RowLaporan In DS.Tbl_VAT.Rows
                        With RowLaporan
                            NoFakturDummy = RowLaporan.Tax_Invoice_No_Referensi
                            NoFakturDummy = NoFakturDummy.Trim
                            NoFakturDummy = NoFakturDummy.Replace("-", "")
                            NoFakturDummy = NoFakturDummy.Replace(".", "")
                            NoNPWP = RowLaporan.NPWP
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
                            + RowLaporan.No_Nota_Retur_Pajak + ";" + Format(CDate(._Date), "dd/MM/yyyy") + ";" + CDate(._Date).Month.ToString + ";" + CDate(._Date).Year.ToString + ";" _
                            + CStr(Math.Floor(CDbl(.Purchase_Sales))) + ";" + CStr(Math.Floor(CDbl(.VAT))) + ";0"

                        End With
                    Next
                    My.Computer.FileSystem.WriteAllText(Me.SaveFileDialog1.FileName, ValTxt, False)
                    MsgBox("Csv file created. Please be aware that characters used in import process (;,"") will be omitted.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KurangkanItemReturn()
        Dim Filter As String = ""
        Dim row1 As DseLaporanVAT1.Item_InvoiceRow
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
End Class