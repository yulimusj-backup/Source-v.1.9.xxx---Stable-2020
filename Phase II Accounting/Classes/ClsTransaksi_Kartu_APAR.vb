Option Explicit On
Option Strict Off

Public Class ClsTransaksi_Kartu_APAR
    Public Enum EnumJenisTransaksiAsal As Integer
        Invoice_Otomatis
        Invoice_Manual
        Invoice_Manual_Transfer
        AOP_Otomatis
        AOP_Manual
        AOP_Manual_Transfer
        SR_Otomatis
        SR_Manual
        SR_Manual_Transfer
        PR_Otomatis
        PR_Manual
        PR_Manual_Transfer
        Sales_Receivement
        Sales_Return_Payment
        Purchase_Payment
        Purchase_Return_Receivement
        Revaluation
        Credit_Note
        Debit_Note
        Invoice_DP_Manual 'add aul
        AOP_DP_Manual 'add aul
    End Enum

    Public Enum EnumJenisKartu As Integer
        AR
        AP
    End Enum

    Public Function Simpan(ByVal vPrimaryTransaksiAsal As Double, ByVal vJenisTransaksi As EnumJenisTransaksiAsal, ByVal vDocumentNo As String, _
     ByVal vPrimaryPihakAcuan As Double, ByVal vDebet As Double, ByVal vKredit As Double, ByVal vDescription As String, _
     ByVal vInOut As EnumJenisKartu, ByVal vDate As Date, ByVal DTA As DseTrans_Kartu_APARTableAdapters.tbltransaksi_kartu_ar_apTableAdapter, _
     ByVal vPrimaryTransaksiAcuan As Double, ByVal vTanggalTransaksiAcuan As Date, ByVal vPaymentTermsTransaksiAcuan As Double, _
     ByVal vJenisTransaksiAcuan As EnumJenisTransaksiAsal, Optional ByVal vPrimaryCurrency As Double = 1, Optional ByVal vDebetOriginal As Decimal = 0, _
    Optional ByVal vCreditOriginal As Decimal = 0, Optional ByVal vExchangeRate As Decimal = 1, Optional ByVal vPrimaryAkun As Double = -1) As Boolean
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Insert(vPrimaryTransaksiAsal, vJenisTransaksi, vPrimaryPihakAcuan, _
            vDocumentNo, vDebet, vKredit, vDescription, vInOut, vDate, vPrimaryTransaksiAcuan, _
            vTanggalTransaksiAcuan, vPaymentTermsTransaksiAcuan, vJenisTransaksiAcuan, vPrimaryCurrency, _
            vExchangeRate, vDebetOriginal, vCreditOriginal, vPrimaryAkun) > 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Simpan(ByVal vPrimaryTransaksiAsal As Double, ByVal vJenisTransaksi As EnumJenisTransaksiAsal, ByVal vDocumentNo As String, _
    ByVal vPrimaryPihakAcuan As Double, ByVal vDebet As Double, ByVal vKredit As Double, ByVal vDescription As String, _
    ByVal vInOut As EnumJenisKartu, ByVal vDate As Date, _
    ByVal vPrimaryTransaksiAcuan As Double, ByVal vTanggalTransaksiAcuan As Date, ByVal vPaymentTermsTransaksiAcuan As Double, _
    ByVal vJenisTransaksiAcuan As EnumJenisTransaksiAsal, Optional ByVal vPrimaryCurrency As Double = 1, Optional ByVal vDebetOriginal As Decimal = 0, _
    Optional ByVal vCreditOriginal As Decimal = 0, Optional ByVal vExchangeRate As Decimal = 1, Optional ByVal vPrimaryAkun As Double = -1) As Boolean
        Dim DTA As New DseTrans_Kartu_APARTableAdapters.tbltransaksi_kartu_ar_apTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Insert(vPrimaryTransaksiAsal, vJenisTransaksi, vPrimaryPihakAcuan, _
            vDocumentNo, vDebet, vKredit, vDescription, vInOut, vDate, vPrimaryTransaksiAcuan, _
            vTanggalTransaksiAcuan, vPaymentTermsTransaksiAcuan, vJenisTransaksiAcuan, vPrimaryCurrency, _
            vExchangeRate, vDebetOriginal, vCreditOriginal, vPrimaryAkun) > 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksiAsal As Double, ByVal vJenisTransaksiAsal As EnumJenisTransaksiAsal) As Boolean
        Dim DTA As New DseTrans_Kartu_APARTableAdapters.tbltransaksi_kartu_ar_apTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Delete(vPrimaryTransaksiAsal, vJenisTransaksiAsal) > 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal vPrimaryCust As Double, ByVal vStartDate As Date, ByVal vEndDate As Date, _
   ByVal vDGV As DataGridView, ByVal TBL As DseTrans_Kartu_APAR.tbltransaksi_kartu_ar_ap1DataTable, _
   ByVal vTipeKartu As EnumJenisKartu, ByVal vDef_Prec As Integer, Optional ByVal vUseSupplierInvoiceNo As Boolean = False) As Boolean
        Dim DTA1 As New DseTrans_Kartu_APARTableAdapters.tbltransaksi_kartu_ar_ap1TableAdapter
        Dim StartingBalance As Double = Me.GetSaldo(vPrimaryCust, vStartDate - TimeSpan.FromDays(1), vTipeKartu)
        Dim row1, row2 As DseTrans_Kartu_APAR.tbltransaksi_kartu_ar_ap1Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim TipeTransaksiAsal As EnumJenisTransaksiAsal = EnumJenisTransaksiAsal.AOP_Otomatis
        Dim DTA_AOP As New DseTrans_Kartu_APARTableAdapters.tbltrans_aopTableAdapter
        Dim TBL_AOP As New DseTrans_Kartu_APAR.tbltrans_aopDataTable
        Dim row_AOP As DseTrans_Kartu_APAR.tbltrans_aopRow

        Dim DTA_CDNote As New DseTrans_Kartu_APARTableAdapters.tbltrans_credit_noteTableAdapter
        Dim TBL_CDNote As New DseTrans_Kartu_APAR.tbltrans_credit_noteDataTable
        Dim row_CDNote As DseTrans_Kartu_APAR.tbltrans_credit_noteRow

        Dim TipeTransaksiAsal2 As EnumJenisTransaksiAsal = EnumJenisTransaksiAsal.Purchase_Payment
        Dim TipeTransaksiAsal3 As EnumJenisTransaksiAsal = EnumJenisTransaksiAsal.Credit_Note
        Dim TipeTransaksiAsal4 As EnumJenisTransaksiAsal = EnumJenisTransaksiAsal.Debit_Note
        Dim TipeTransaksiAsal5 As EnumJenisTransaksiAsal = EnumJenisTransaksiAsal.AOP_DP_Manual 'aul add

        'aul add
        Dim DtaManualDP As New DseTrans_Kartu_APARTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim TblManualDP As New DseTrans_Kartu_APAR.tbltrans_historicalbalancemanual_dpDataTable
        Dim rowManualDP As DseTrans_Kartu_APAR.tbltrans_historicalbalancemanual_dpRow

        Try
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA_AOP.Connection.ConnectionString = DTA1.Connection.ConnectionString
            DTA_CDNote.Connection.ConnectionString = DTA1.Connection.ConnectionString
            DtaManualDP.Connection.ConnectionString = DTA1.Connection.ConnectionString 'aul add

            DTA1.Fill(TBL, vPrimaryCust, vTipeKartu, vEndDate, vStartDate)
            'ganti no transaksi, jika tipenya kartu ap
            If vTipeKartu = EnumJenisKartu.AP And TBL.Rows.Count > 0 Then
                If vUseSupplierInvoiceNo = True Then
                    DTA_AOP.FillBy_All(TBL_AOP)
                    DTA_CDNote.Fill(TBL_CDNote)
                    DtaManualDP.Fill(TblManualDP) 'aul add

                    For Each row1 In TBL.Rows

                        If row1.Type_Transaksi_Asal = TipeTransaksiAsal Then
                            TBL_AOP.DefaultView.RowFilter = "Primary = " + row1.Primary_Transaksi_Asal.ToString
                            row_AOP = TBL_AOP.DefaultView(0).Row
                            row1.Document_No = row_AOP.Supplier_Invoice_No
                        End If

                        If row1.Type_Transaksi_Asal = TipeTransaksiAsal2 Then
                            TBL_AOP.DefaultView.RowFilter = "[No Transaksi] = '" + row1.Remarks + "'"
                            If TBL_AOP.DefaultView.Count > 0 Then
                                row_AOP = TBL_AOP.DefaultView(0).Row
                                row1.Remarks = row_AOP.Supplier_Invoice_No
                            Else
                                TBL_CDNote.DefaultView.RowFilter = "[No Transksi] = '" + row1.Remarks + "'"
                                If TBL_CDNote.DefaultView.Count > 0 Then
                                    row_CDNote = TBL_CDNote.DefaultView(0).Row
                                    row1.Remarks = row_CDNote.customer_debet_note
                                End If
                            End If
                        End If

                        If row1.Type_Transaksi_Asal = TipeTransaksiAsal3 OrElse row1.Type_Transaksi_Asal = TipeTransaksiAsal4 Then
                            TBL_CDNote.DefaultView.RowFilter = "[Primary] = " & row1.Primary_Transaksi_Asal
                            If TBL_CDNote.DefaultView.Count > 0 Then
                                row_CDNote = TBL_CDNote.DefaultView(0).Row
                                row1.Document_No = row_CDNote.customer_debet_note
                            End If
                        End If

                        'aul add
                        If row1.Type_Transaksi_Asal = TipeTransaksiAsal5 Then
                            TblManualDP.DefaultView.RowFilter = "[Primary] = " & row1.Primary_Transaksi_Asal
                            If TblManualDP.DefaultView.Count > 0 Then
                                rowManualDP = TblManualDP.DefaultView(0).Row
                                row1.Document_No = rowManualDP.Invoice_No
                            End If
                        End If
                    Next
                End If
            End If





            'masukkan starting balance
            'If TBL.Rows.Count = 0 Then
            '    TBL.Rows.Add(vStartDate, "-", "Starting Balance", 0, 0, StartingBalance, 0, -1)
            'Else
            Dim row As DseTrans_Kartu_APAR.tbltransaksi_kartu_ar_ap1Row = TBL.NewRow
            row._Date = vStartDate
            row.Balance = StartingBalance
            row.Document_No = "-"
            row.Debet = 0
            row.Remarks = "Starting Balance"
            row.Kredit = 0
            row.Type_Transaksi_Asal = 0
            row.Primary_Transaksi_Asal = -1

            TBL.Rows.InsertAt(row, 0)
            'End If

            'hitung balance untuk setiap itemnya
            For i As Integer = 0 To TBL.Rows.Count - 1
                row1 = TBL.Rows(i)
                If i <> 0 Then
                    row2 = TBL.Rows(i - 1)

                    If vTipeKartu = EnumJenisKartu.AR Then
                        row1.Balance = row2.Balance + row1.Debet - row1.Kredit
                    Else
                        row1.Balance = row2.Balance + row1.Kredit - row1.Debet
                    End If
                End If

                'format data
                row1.Balance_Str = ObjCD.DigitGrouping(row1.Balance, vDef_Prec)
                row1.Debet_Str = ObjCD.DigitGrouping(row1.Debet, vDef_Prec)
                row1.Kredit_Str = ObjCD.DigitGrouping(row1.Kredit, vDef_Prec)
            Next

            vDGV.DataSource = TBL.DefaultView
            vDGV.Columns(0).DefaultCellStyle.Format = "dd MMM yyyy"
            vDGV.Columns(3).Visible = False
            vDGV.Columns(4).Visible = False
            vDGV.Columns(5).Visible = False
            vDGV.Columns(2).HeaderText = "Reference"
            vDGV.Columns(6).HeaderText = "Debet"
            vDGV.Columns(7).HeaderText = "Credit"
            vDGV.Columns(8).HeaderText = "Balance"
            vDGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            vDGV.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            vDGV.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            vDGV.Columns(9).Visible = False
            vDGV.Columns(10).Visible = False

            Dim col As DataGridViewColumn
            For Each col In vDGV.Columns
                col.SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            vDGV.Refresh()
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            ObjCD = Nothing
            DTA_AOP = Nothing
            TBL_AOP = Nothing
        End Try
    End Function

    Public Function GetSaldo(ByVal vPrimaryPihakAcuan As Double, ByVal Tanggal As Date, ByVal vTipe As EnumJenisKartu) As Double
        Dim TotalDebet, TotalKredit As Double
        Dim DTA As New DseTrans_Kartu_APARTableAdapters.tbltransaksi_kartu_ar_apTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetTotalDebet(vPrimaryPihakAcuan, vTipe, Tanggal).HasValue Then
                TotalDebet = DTA.ScalarQuery_GetTotalDebet(vPrimaryPihakAcuan, vTipe, Tanggal).Value
            Else
                TotalDebet = 0
            End If
            If DTA.ScalarQuery_GetTotalCredit(vPrimaryPihakAcuan, vTipe, Tanggal).HasValue Then
                TotalKredit = DTA.ScalarQuery_GetTotalCredit(vPrimaryPihakAcuan, vTipe, Tanggal).Value
            Else
                TotalKredit = 0
            End If

            If vTipe = EnumJenisKartu.AP Then
                Return TotalKredit - TotalDebet
            Else
                Return TotalDebet - TotalKredit
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipeTransaksiAsalKartuAPAR(ByVal vTipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance) As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal
        Try
            If vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual_Transfer
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note
            ElseIf vTipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note Then
                Return ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetTipeSalesBalance(ByVal vTipeTransaksiAsal_KartuAPAR As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal) As ClsTransaksiSalesBalance.EnumTipeSalesBalance
        Try
            If vTipeTransaksiAsal_KartuAPAR = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual Or _
            vTipeTransaksiAsal_KartuAPAR = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual Then
                Return ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
            ElseIf vTipeTransaksiAsal_KartuAPAR = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual_Transfer Or _
            vTipeTransaksiAsal_KartuAPAR = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual_Transfer Then
                Return ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_Transfer_Invoice
            ElseIf vTipeTransaksiAsal_KartuAPAR = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis Or _
            vTipeTransaksiAsal_KartuAPAR = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis Then
                Return ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
            ElseIf vTipeTransaksiAsal_KartuAPAR = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual Or _
            vTipeTransaksiAsal_KartuAPAR = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual Then
                Return ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
            ElseIf vTipeTransaksiAsal_KartuAPAR = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis Or _
            vTipeTransaksiAsal_KartuAPAR = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis Then
                Return ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR
            ElseIf vTipeTransaksiAsal_KartuAPAR = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Then
                Return ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
            ElseIf vTipeTransaksiAsal_KartuAPAR = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note Then
                Return ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
