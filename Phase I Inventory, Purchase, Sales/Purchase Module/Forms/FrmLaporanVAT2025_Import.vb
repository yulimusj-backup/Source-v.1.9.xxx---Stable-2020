Option Explicit On
Option Strict Off

Imports System.Data
Imports Microsoft.VisualBasic.FileIO

Public Class FrmLaporanVAT2025_Import
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public tblDummy As New DseLaporanVAT2025.Tbl_VATDataTable
    Dim tblDummy1 As New DseLaporanVAT2025.Tbl_VATDataTable
    Dim tblImport As New DseLaporanVAT2025_1.TblImportTaxNumberDataTable
    Dim dataValid As Integer = 0
    Dim objCD As New Core_Module.ClsCekData

    Public Enum EnumColumnImportTax As Integer
        NPWP
        NPWP_Name
        Transaction_Code
        Tax_No
        Tax_Date
        Tax_Month
        Tax_Year
        Tax_Status
        Esign
        Tax_base
        Other_Tax_base
        VAT
        PPnBM
        Signed_By
        Reference_No
        Report_By_Seller
        Report_By_VAT_Collector
    End Enum

    Public Enum EnumColumnImportTax2 As Integer
        Stringkosong
        NPWP
        NPWP_Name
        Transaction_Code
        Tax_No
        Tax_Date
        Tax_Month
        Tax_Year
        Tax_Status
        Esign
        Tax_base
        Other_Tax_base
        VAT
        PPnBM
        Signed_By
        Reference_No
        Report_By_Seller
        Report_By_VAT_Collector
        Stringkosong2
    End Enum

    Private Sub FrmImportData_MasterCustomer_Invalid_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            ' x min = 100, x max = 535
            ' y min = 639, y max = 1200
            Me.Width = 639
            Me.Height = 100
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            ' Set the filter to allow only CSV files
            OpenFileDialog1.Filter = "CSV Files (*.csv)|*.csv"
            ' Optionally, set other properties
            OpenFileDialog1.Title = "Select a CSV File"
            OpenFileDialog1.Multiselect = False ' Allow only single file selection
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                Me.BackgroundWorker1.RunWorkerAsync()
                'Me.Close()
            Else
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Private Function LoadCSV(ByVal vFilePath As String) As Boolean
        Dim StreamReader_ As System.IO.StreamReader
        Dim CSV_Line As String = ""
        Dim Row As DseLaporanVAT2025_1.TblImportTaxNumberRow
        Dim counter As Integer = 0
        Dim Separator As String = ","
        Try
            StreamReader_ = New System.IO.StreamReader(vFilePath)
            CSV_Line = StreamReader_.ReadLine()
            Me.tblImport.Rows.Clear()
            Do
                counter += 1
                CSV_Line = StreamReader_.ReadLine
                If Not CSV_Line = String.Empty Then
                    Row = Me.tblImport.NewTblImportTaxNumberRow
                    Dim CSV_Column() As String = CSV_Line.Split(Separator)
                    'Kondisi CSV langsung CoreTax sekarang baris judul terdapat 2 kolom pertama kosong dan baris lainnya kolom pertama kosong
                    If String.IsNullOrEmpty(CSV_Column(EnumColumnImportTax.NPWP)) = True Then
                        'baris ke-2 ambil enum EnumColumnImportTax2
                        Row.No_NPWP = CSV_Column(EnumColumnImportTax2.NPWP).Trim.ToUpper.Replace("'", "")
                        Row.NPWP_Name = CSV_Column(EnumColumnImportTax2.NPWP_Name).Trim.ToUpper.Replace("'", "")
                        Row.Kode_Transaksi = CSV_Column(EnumColumnImportTax2.Transaction_Code).Trim.ToUpper.Replace("'", "")
                        Row.Nomor_Faktur_Pajak = CSV_Column(EnumColumnImportTax2.Tax_No).Trim.ToUpper.Replace("'", "")
                        Row.Tanggal_Faktur_Pajak = Format(CDate(CSV_Column(EnumColumnImportTax2.Tax_Date)), "yyyy-MM-dd")
                        Row.Masa_Pajak = CSV_Column(EnumColumnImportTax2.Tax_Month).Trim.ToUpper.Replace("'", "")
                        Row.Tahun = CSV_Column(EnumColumnImportTax2.Tax_Year).Trim.ToUpper.Replace("'", "")
                        Row.Status_Faktur = CSV_Column(EnumColumnImportTax2.Tax_Status).Trim.ToUpper.Replace("'", "")
                        Row.ESign = CSV_Column(EnumColumnImportTax2.Esign).Trim.ToUpper.Replace("'", "")
                        Row.DPP = CDec(CSV_Column(EnumColumnImportTax2.Tax_base).Trim.ToUpper.Replace("'", ""))
                        Row.DPP_Nilai_Lain = CDec(CSV_Column(EnumColumnImportTax2.Other_Tax_base).Trim.ToUpper.Replace("'", ""))
                        Row.PPN = CDec(CSV_Column(EnumColumnImportTax2.VAT).Trim.ToUpper.Replace("'", ""))
                        Row.PPnBM = CDec(CSV_Column(EnumColumnImportTax2.PPnBM).Trim.ToUpper.Replace("'", ""))
                        Row.Penandatangan = CSV_Column(EnumColumnImportTax2.Signed_By).Trim.ToUpper.Replace("'", "")
                        Row.Referensi = CSV_Column(EnumColumnImportTax2.Reference_No).Trim.ToUpper.Replace("'", "")
                        Row.Dilaporkan_Oleh_Penjual = CSV_Column(EnumColumnImportTax2.Report_By_Seller).Trim.ToUpper.Replace("'", "")
                        Row.Dilaporkan_Oleh_Pemungut_PPN = CSV_Column(EnumColumnImportTax2.Report_By_VAT_Collector).Trim.ToUpper.Replace("'", "")
                    Else
                        'jika CSV tidak ada baris kosong (dibuat dari excel save as CSV), maka pakai enum EnumColumnImportTax
                        Row.No_NPWP = CSV_Column(EnumColumnImportTax.NPWP).Trim.ToUpper.Replace("'", "")
                        Row.NPWP_Name = CSV_Column(EnumColumnImportTax.NPWP_Name).Trim.ToUpper.Replace("'", "")
                        Row.Kode_Transaksi = CSV_Column(EnumColumnImportTax.Transaction_Code).Trim.ToUpper.Replace("'", "")
                        Row.Nomor_Faktur_Pajak = CSV_Column(EnumColumnImportTax.Tax_No).Trim.ToUpper.Replace("'", "")
                        Row.Tanggal_Faktur_Pajak = Format(CDate(CSV_Column(EnumColumnImportTax.Tax_Date)), "yyyy-MM-dd")
                        Row.Masa_Pajak = CSV_Column(EnumColumnImportTax.Tax_Month).Trim.ToUpper.Replace("'", "")
                        Row.Tahun = CSV_Column(EnumColumnImportTax.Tax_Year).Trim.ToUpper.Replace("'", "")
                        Row.Status_Faktur = CSV_Column(EnumColumnImportTax.Tax_Status).Trim.ToUpper.Replace("'", "")
                        Row.ESign = CSV_Column(EnumColumnImportTax.Esign).Trim.ToUpper.Replace("'", "")
                        Row.DPP = CDec(CSV_Column(EnumColumnImportTax.Tax_base).Trim.ToUpper.Replace("'", ""))
                        Row.DPP_Nilai_Lain = CDec(CSV_Column(EnumColumnImportTax.Other_Tax_base).Trim.ToUpper.Replace("'", ""))
                        Row.PPN = CDec(CSV_Column(EnumColumnImportTax.VAT).Trim.ToUpper.Replace("'", ""))
                        Row.PPnBM = CDec(CSV_Column(EnumColumnImportTax.PPnBM).Trim.ToUpper.Replace("'", ""))
                        Row.Penandatangan = CSV_Column(EnumColumnImportTax.Signed_By).Trim.ToUpper.Replace("'", "")
                        Row.Referensi = CSV_Column(EnumColumnImportTax.Reference_No).Trim.ToUpper.Replace("'", "")
                        Row.Dilaporkan_Oleh_Penjual = CSV_Column(EnumColumnImportTax.Report_By_Seller).Trim.ToUpper.Replace("'", "")
                        Row.Dilaporkan_Oleh_Pemungut_PPN = CSV_Column(EnumColumnImportTax.Report_By_VAT_Collector).Trim.ToUpper.Replace("'", "")
                    End If
                    Me.tblImport.Rows.Add(Row)
                Else
                    Exit Do
                End If
            Loop
            Return True
        Catch ex As Exception
            Me.tblImport.Rows.Clear()
            If ex.Message.ToLower.Contains("because it is being used by another process") Then
                MsgBox("Please close the CSV file that opened by another program.", MsgBoxStyle.Information, "Information")
            Else
                MsgBox("Something wrong while load CSV file: " & vbNewLine & ex.Message & vbNewLine & "Please contact your software provider.", MsgBoxStyle.Information, "Information")
            End If
        Finally
            Try
                StreamReader_.Close()
            Catch ex As Exception
            End Try
        End Try
    End Function

    Function CheckReference(ByVal vRowIndexFrom As Integer, ByVal vRowIndexTo As Integer, ByVal vReloadReferenceData As Boolean, Optional ByVal vCalledFromCellEndEdit As Boolean = False) As Boolean
        Dim RowDgv As DataGridViewRow
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.Cursor = Cursors.Default
        End Try
    End Function

    Private Function FilterTaxNo() As Integer
        Dim row As DseLaporanVAT2025_1.TblImportTaxNumberRow
        Dim rowDummy As DseLaporanVAT2025.Tbl_VATRow
        Dim rowDummy1 As DseLaporanVAT2025.Tbl_VATRow
        Dim tblCopy As New DseLaporanVAT2025_1.TblImportTaxNumberDataTable
        Dim InvalidStatus As String = "Invoice status must be approved"
        Dim InvalidNoFakturKosong As String = "Invoice number cannot be empty"
        Dim InvalidFormatNoFaktur As String = "Invoice number must be 17 digits"
        Dim InvalidReference As String = "Reference not found"
        Dim CompareReference As Boolean
        Dim Count As Integer = 1
        Try
            Me.ProgressBar1.Minimum = 1
            Me.ProgressBar1.Maximum = Me.tblImport.Count
            tblCopy = Me.tblImport.Copy
            Me.tblDummy1.Rows.Clear()
            For Each row In Me.tblImport.Rows
                Me.ProgressBar1.Value = Count
                CompareReference = True
                'Cek Status Faktur
                If row.Status_Faktur <> "APPROVED" Then
                    If row.Invalid_Reason = "" Then
                        row.Invalid_Reason = InvalidStatus
                    Else
                        row.Invalid_Reason = row.Invalid_Reason & ", " & InvalidStatus
                    End If
                    CompareReference = False
                End If

                'Cek No Faktur Kosong
                If String.IsNullOrEmpty(row.Nomor_Faktur_Pajak.ToString.Trim) = True Then
                    If row.Invalid_Reason = "" Then
                        row.Invalid_Reason = InvalidNoFakturKosong
                    Else
                        row.Invalid_Reason = row.Invalid_Reason & ", " & InvalidNoFakturKosong
                    End If
                    CompareReference = False
                Else
                    'Cek format no faktur adalah 17 digit
                    If row.Nomor_Faktur_Pajak.ToString.Trim.Length <> 17 Then
                        If row.Invalid_Reason = "" Then
                            row.Invalid_Reason = InvalidFormatNoFaktur
                        Else
                            row.Invalid_Reason = row.Invalid_Reason & ", " & InvalidFormatNoFaktur
                        End If
                        CompareReference = False
                    End If
                End If

                'ketika format faktur benar dan approved, compare reference
                Dim referenceString As String = ""
                If row.Referensi.StartsWith("INVOICE NO: ") = True Then
                    referenceString = row.Referensi.Substring("INVOICE NO: ".Length - 1).Trim
                Else
                    referenceString = row.Referensi.Trim
                End If

                'ketika no referensi kosong
                If String.IsNullOrEmpty(referenceString) Then
                    CompareReference = False
                    If String.IsNullOrEmpty(row.Invalid_Reason) Then
                        row.Invalid_Reason = InvalidReference
                    Else
                        row.Invalid_Reason = row.Invalid_Reason & ", " & InvalidReference
                    End If
                End If

                If CompareReference = True Then
                    'Cek apakah ada Reference yang sama dengan TblMaster
                    Me.tblDummy.DefaultView.RowFilter = ""
                    Me.tblDummy.DefaultView.RowFilter = "[Invoice_No] LIKE '" & referenceString & "%'"
                    Dim dv As DataRowView
                    If tblDummy.DefaultView.Count > 0 Then
                        For Each dv In Me.tblDummy.DefaultView
                            rowDummy = dv.Row
                            rowDummy.kode_jenis_transaksi_pajak = row.Nomor_Faktur_Pajak.Substring(0, 4) 'substring 4 awal ambil kode transaksi (ex. 0100)
                            rowDummy.Tax_Invoice_No = row.Nomor_Faktur_Pajak.Substring(4) 'substring 4 tidak ambil kode transaksi (ex. 0100)
                            rowDummy1 = Me.tblDummy1.NewTbl_VATRow
                            rowDummy1.ItemArray = rowDummy.ItemArray
                            Me.tblDummy1.Rows.Add(rowDummy1)
                        Next
                    Else
                        If row.Invalid_Reason = "" Then
                            row.Invalid_Reason = InvalidReference
                        Else
                            row.Invalid_Reason = row.Invalid_Reason & ", " & InvalidReference
                        End If
                    End If
                End If
                Me.lblProgress.Text = "Progress " & Me.ProgressBar1.Value & " of " & Me.ProgressBar1.Maximum & " data"
                Count += 1
            Next

            Me.lblProgress.Text = Me.ProgressBar1.Value & " of " & Me.ProgressBar1.Maximum & " data have been successfully processed"
            Me.tblImport.DefaultView.RowFilter = "[Invalid Reason]=''"
            Return Me.tblImport.DefaultView.Count
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(10)
            If Me.LoadCSV(Me.OpenFileDialog1.FileName) Then
                Me.TxtFilePath.Text = Me.OpenFileDialog1.FileName
                dataValid = Me.FilterTaxNo()
            Else
                Me.TxtFilePath.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        'SEMENTARA
        Dim TotalInvalid As Integer = 0
        Dim TotalValid As Integer = 0
        Dim dv As DataRowView
        Dim dvDummy As DataRowView
        Dim row As DseLaporanVAT2025_1.TblImportTaxNumberRow
        Dim rowDummy As DseLaporanVAT2025.Tbl_VATRow
        'Dim rowMaster As DseTrans_CreateEFaktur2025_Cari.tbl_masterRow
        Dim rowdgv As DataGridViewRow
        Dim DTA As New DseLaporanVAT2025_1TableAdapters.tblInvoice_InvoiceDPTableAdapter

        Try
            ' x min = 100, x max = 535
            ' y min = 639, y max = 1200
            'copy dari tblImport ke dgv jika ada yang invalid
            'DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.tblImport.DefaultView.RowFilter = "[Invalid Reason]<>''"
            TotalInvalid = Me.tblImport.DefaultView.Count
            If TotalInvalid > 0 Then 'input dgv
                Me.Width = 1200
                Me.Height = 535
                Me.MinimumSize = New Size(1200, 535)
                Me.ControlBox = True
                Me.GroupBox1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
                Me.BtnClose.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
                Me.CenterToScreen()
                For Each dv In Me.tblImport.DefaultView
                    row = dv.Row
                    rowdgv = Me.DgvItem.Rows(Me.DgvItem.Rows.Add)
                    rowdgv.Cells(Me.ColNPWP.Index).Value = row.No_NPWP
                    rowdgv.Cells(Me.ColCustomerName.Index).Value = row.NPWP_Name
                    rowdgv.Cells(Me.ColTransactionCode.Index).Value = row.Kode_Transaksi
                    rowdgv.Cells(Me.ColTaxInvoiceNumber.Index).Value = row.Nomor_Faktur_Pajak
                    rowdgv.Cells(Me.ColTaxDate.Index).Value = Format(CDate(row.Tanggal_Faktur_Pajak), "yyyy-MM-dd")
                    rowdgv.Cells(Me.ColTaxPeriod.Index).Value = row.Masa_Pajak
                    rowdgv.Cells(Me.ColTaxYear.Index).Value = row.Tahun
                    rowdgv.Cells(Me.ColTaxStatus.Index).Value = row.Status_Faktur
                    rowdgv.Cells(Me.ColTaxBase.Index).Value = CDec(row.DPP)
                    rowdgv.Cells(Me.ColOtherTaxBase.Index).Value = CDec(row.DPP_Nilai_Lain)
                    rowdgv.Cells(Me.ColVAT.Index).Value = CDec(row.PPN)
                    rowdgv.Cells(Me.ColPPnBM.Index).Value = CDec(row.PPnBM)
                    rowdgv.Cells(Me.ColSignedBy.Index).Value = row.Penandatangan
                    rowdgv.Cells(Me.ColReference.Index).Value = row.Referensi
                    rowdgv.Cells(Me.ColReportedBySeller.Index).Value = row.Dilaporkan_Oleh_Penjual
                    rowdgv.Cells(Me.ColReportedByVATCollector.Index).Value = row.Dilaporkan_Oleh_Pemungut_PPN
                    rowdgv.Cells(Me.ColInvalidReason.Index).Value = row.Invalid_Reason
                Next
                Me.objCD.SetNumber(Me.DgvItem)
            End If

            Me.tblImport.DefaultView.RowFilter = "[Invalid Reason]=''"
            TotalValid = Me.tblImport.DefaultView.Count
            If TotalValid > 0 Then 'update table custom sales invoice
                'For Each dv In Me.tblImport.DefaultView
                '    row = dv.Row
                '    Me.tblDummy1.DefaultView.RowFilter = ""
                '    Me.tblDummy1.DefaultView.RowFilter = "[Invoice No] LIKE '" & row.Referensi & "%'"
                '    If Me.tblDummy1.DefaultView.Count > 0 Then
                '        For Each dvDummy In tblDummy1.DefaultView
                '            rowDummy = dvDummy.Row
                '            DTA.UpdateQuery_NoPajakBy_PrimaryInvoice(-1, row.Kode_Transaksi, row.Nomor_Faktur_Pajak, rowDummy.primary_sales_invoice)
                '        Next
                '    End If
                'Next
                For Each rowDummy In Me.tblDummy1
                    With rowDummy
                        DTA.UpdateQuery_NoPajakBy_PrimaryInvoice(-1, .kode_jenis_transaksi_pajak, .Tax_Invoice_No, .primary_sales_invoice)
                    End With
                Next
            End If

            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.Cursor = Cursors.Default
            MsgBox(TotalValid & " Data Imported.", MsgBoxStyle.Information, "Information")
            If TotalInvalid = 0 Then
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.Sorted
        Try
            Me.objCD.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class